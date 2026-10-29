"""Revision-bound query coordination; business matching lives in customer_allocation."""
import json
import threading
import uuid


class CustomerStateError(ValueError):
    pass


class CustomerService:
    def __init__(self, book_factory, lock):
        self.book_factory = book_factory
        self.lock = lock

    def assert_idle(self, session):
        if session.get('customerExport', {}).get('status') == 'running':
            raise CustomerStateError('统计报表正在导出，请完成后再修改策略、客户表或品规范围。')

    def invalidate(self, session, strategy=False):
        self.assert_idle(session)
        session['customerEpoch'] = session.get('customerEpoch', 0) + 1
        session.pop('customerQueries', None)
        session.pop('customerEngine', None)
        if strategy:
            session.pop('confirmedChoices', None)
            session.pop('confirmedChoicesToken', None)

    def confirm_choices(self, session, choices, token):
        with self.lock:
            self.assert_idle(session)
            expected = {str(p['sourceRow']) for p in session['review']['exportProducts']}
            if (token != session['review']['zeroToken'] or not isinstance(choices, dict)
                    or set(choices) != expected or any(v not in ('keep', 'omit') for v in choices.values())):
                raise CustomerStateError('品规清单已变化，请重新确认。')
            if 'keep' not in choices.values():
                raise CustomerStateError('请至少保留一个参与测算的品规。')
            if session.get('confirmedChoices') != choices or session.get('confirmedChoicesToken') != token:
                self.invalidate(session)
                session['confirmedChoices'] = choices.copy()
                session['confirmedChoicesToken'] = token
            return self.state(session)

    def upload(self, session, content, filename, expected_token):
        with self.lock:
            self.assert_idle(session)
            if expected_token != session['review']['zeroToken']:
                raise CustomerStateError('策略已变化，请重新上传客户表。')
            template = session['templateKey']
            epoch = session.get('customerEpoch', 0)
        try:
            from .customer_allocation import parse_customers
        except ImportError:
            from customer_allocation import parse_customers
        dataset = parse_customers(content, template, self.book_factory)
        with self.lock:
            self.assert_idle(session)
            if expected_token != session['review']['zeroToken'] or epoch != session.get('customerEpoch', 0):
                raise CustomerStateError('上传期间数据已变化，请重新上传客户表。')
            self.invalidate(session)
            session['customerDataset'] = dataset
            session['customerFilename'] = filename
            return self.state(session)

    def state(self, session):
        dataset = session.get('customerDataset')
        confirmed = session.get('confirmedChoicesToken') == session['review']['zeroToken'] and bool(session.get('confirmedChoices'))
        return {
            'ok': True, 'revision': session.get('customerEpoch', 0),
            'reviewToken': session['review']['zeroToken'], 'filename': session.get('customerFilename', ''),
            'customerSummary': (dataset.summary if hasattr(dataset, 'summary') else dataset.get('summary', {})) if dataset is not None else None,
            'choicesConfirmed': confirmed, 'choices': session.get('confirmedChoices') if confirmed else None,
            'exportProducts': session['review']['exportProducts'],
            'busy': session.get('customerExport', {}).get('status') == 'running',
        }

    def effective_choices(self, session):
        if self.state(session)['choicesConfirmed']:
            return session['confirmedChoices']
        return {str(p['sourceRow']): p.get('defaultChoice', 'keep')
                for p in session['review']['exportProducts']}

    def _engine(self, session):
        if 'customerDataset' not in session:
            raise CustomerStateError('请先上传客户基本信息表。')
        if 'customerEngine' not in session:
            try:
                from .customer_allocation import AllocationQuery
            except ImportError:
                from customer_allocation import AllocationQuery
            session['customerEngine'] = AllocationQuery(session['customerDataset'], session['review'],
                                                       self.effective_choices(session), session['templateKey'])
        return session['customerEngine']

    def options(self, session):
        with self.lock:
            if 'customerDataset' in session:
                options = self._engine(session).filter_options()
            else:
                choices = self.effective_choices(session)
                options = {'products': [dict(
                    sourceRow=p['sourceRow'], productCode=p['productCode'], productName=p['productName'],
                    allocationStatus='已排除投放' if choices[str(p['sourceRow'])] == 'omit' else '参与投放')
                    for p in session['review']['exportProducts']], 'counties': [], 'markets': []}
            return dict(self.state(session), options=options)

    def get_query(self, session, query_id):
        query = session.get('customerQueries', {}).get(query_id)
        if not query or query['epoch'] != session.get('customerEpoch', 0) or query['token'] != session['review']['zeroToken']:
            raise CustomerStateError('查询结果已失效，请重新测算后再操作。')
        return query

    def query(self, session, payload):
        with self.lock:
            if payload.get('queryId'):
                ident = payload['queryId']
                query = self.get_query(session, ident)
            else:
                self.assert_idle(session)
                if payload.get('revision') != session.get('customerEpoch', 0):
                    raise CustomerStateError('客户或品规范围已变化，请刷新数据后重新测算。')
                view = payload.get('view', 'summary')
                filters = payload.get('filters', {})
                if (not isinstance(filters, dict) or filters.get('weekParity') not in ('odd', 'even')
                        or type(filters.get('monthlyWeek')) is not bool):
                    raise CustomerStateError('请先选择单周或双周，并确认是否为每月一订的投放周。')
                plan = self._engine(session).plan_query(filters, view)
                ident = uuid.uuid4().hex
                query = {'plan': plan, 'view': view, 'filters': json.loads(json.dumps(filters)),
                         'epoch': session.get('customerEpoch', 0), 'token': session['review']['zeroToken']}
                cache = session.setdefault('customerQueries', {})
                cache[ident] = query
                while len(cache) > 12:
                    del cache[next(iter(cache))]
            plan = query['plan']
            return dict(plan.page(payload.get('page', 1), payload.get('pageSize', 50)), ok=True,
                        queryId=ident, revision=query['epoch'], columns=plan.columns,
                        totals=plan.totals, view=query['view'])

    def start_export(self, session, payload, saved_files, lifetime):
        with self.lock:
            self.assert_idle(session)
            query = self.get_query(session, payload.get('queryId'))
            save_id, session_id = payload.get('saveId'), payload.get('sessionId')
            target = saved_files.target(save_id, session_id)
            if target.suffix.lower() != '.xlsx':
                raise CustomerStateError('统计报表必须保存为 .xlsx 文件。')
            plan = query['plan']
            if plan.row_count == 0:
                raise CustomerStateError('当前查询没有可导出的记录。')
            job = {'id': uuid.uuid4().hex, 'status': 'running', 'rowCount': 0, 'sheetCount': 0,
                   'totalRows': plan.row_count}
            metadata = self.export_metadata(session, query)
            session['customerExport'] = job
            lifetime.begin_work()

        def work():
            try:
                try:
                    from .query_workbook import write_query_workbook
                except ImportError:
                    from query_workbook import write_query_workbook

                def progress(data):
                    with self.lock:
                        job.update(data)

                def produce(path):
                    return write_query_workbook(path, plan.columns, plan.iter_rows(), metadata, progress=progress)

                path, result = saved_files.write_generated(save_id, session_id, produce)
                with self.lock:
                    job.update(status='complete', result=dict(result, selectedPath=str(path), filename=path.name, saveId=save_id))
            except Exception as exc:
                with self.lock:
                    job.update(status='failed', error='统计导出失败：%s' % exc)
            finally:
                lifetime.end_work()

        thread = threading.Thread(target=work, name='customer-xlsx-export', daemon=False)
        try:
            thread.start()
        except Exception:
            with self.lock:
                job['status'] = 'failed'
            lifetime.end_work()
            raise
        return {'ok': True, 'jobId': job['id'], 'totalRows': job['totalRows']}

    def export_status(self, session, job_id):
        with self.lock:
            job = session.get('customerExport')
            if not job or job['id'] != job_id:
                raise CustomerStateError('导出任务不存在。')
            return dict(job, ok=True)

    def export_metadata(self, session, query):
        filters = query['filters']
        products = {p['sourceRow']: p['productCode']+' '+p['productName'] for p in session['review']['exportProducts']}
        chosen = filters.get('productRows')
        labels = '全部参与测算品规' if chosen is None else '、'.join(products.get(p, str(p)) for p in chosen)
        fmt = lambda values: '全部' if values is None else ('、'.join(values) or '未选择')
        totals = query['plan'].totals
        return [
            {'summary':'品规汇总','detail':'客户明细','issues':'未匹配清单'}[query['view']] + ' - 策略测算投放量（非实际订货量）',
            '客户表：'+session.get('customerFilename', ''),
            '品规范围已复核。' if self.state(session)['choicesConfirmed'] else '品规范围待复核：当前按默认建议统计。',
            '已排除投放的品规仍保留在查询结果中，投放量按0条统计。',
            '品规：'+labels,
            '营销部：'+fmt(filters.get('counties'))+'；市场类型：'+fmt(filters.get('markets')),
            '档位：%s-%s；1 箱 = 250 条；不按周几排除客户。' % (filters.get('minTier', 1), filters.get('maxTier', 30)),
            '策略周：%s；每月一订投放周：%s。' % ('单周' if filters.get('weekParity') == 'odd' else '双周', '是' if filters.get('monthlyWeek') else '否'),
            '订货周期或周次异常客户：%s，已按0条计算，请复核异常清单。' % totals.get('cycleReviewCustomerCount', 0),
            '本周应投客户数（去重）：%s。' % totals.get('eligibleCustomerCount', 0),
            '筛选范围全部客户（含非本周订货客户，去重）：%s；参与品规：%s；测算投放量：%s 条 / %.3f 箱。' % (
                totals.get('customerCount', 0), totals.get('productCount', 0),
                totals.get('quantityTiao', 0), totals.get('quantityXiang', 0)),
            '数据存在缺项，未匹配客户未按0条计入。' if totals.get('incomplete') else '当前筛选范围匹配完整。',
        ]
