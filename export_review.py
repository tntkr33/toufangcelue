"""Export suggestions are separate from quantity parsing and template matching."""
import re


def build_export_products(products, edits, county_names):
    counties = {re.sub(r'[县市]$', '', name.replace('营销部', '')) for name in county_names if name}
    counties = {name for name in counties if len(name) >= 2}
    price_pattern = r'按(?:价位|价格)段投放'
    result = []
    for product in products:
        text = re.sub(r'\s+', '', product['strategyText'])
        rules = product['rules']
        numbers = edits.get(str(product['sourceRow']), {}).get('numbers', {})
        positive = any(v > 0 for row in numbers.values() for v in row.values())
        explicit_stop = re.fullmatch(r'(?:无库存|库存不足|停投|暂停投放|不投放|无货)+',
                                     re.sub(r'[，,。；;、：:]+', '', text))
        whole_stop = bool(explicit_stop and rules and not product['warnings'] and not positive and all(
            r.get('zero') and not r.get('counties') and not r.get('marketTypes') for r in rules))
        has_price = bool(re.search(price_pattern, text))
        global_price = re.match(r'^(?:(?:本周|本期|全市|全部)[：:]?)?' + price_pattern, text)
        scoped = any(name in text for name in counties) or re.search(
            r'营销部|县区|区县|部分|其他|其它|其余|仅|限定|指定|除外|例外|城市|县城|乡镇|农村|市区|城区|镇区|商圈', text)
        uncertain = re.search(r'档|待定|待确认|是否|暂未|待通知|可能|或者', text)
        # Only known global limits/notes may follow the global policy. Unknown
        # trailing instructions can contain an unrecognized scope, so retain them.
        remainder = text[global_price.end():] if global_price else text
        remainder = re.sub(
            r'(?:投放)?(?:上限|下限|保底)\d+条|参考同价位段销量|(?:由)?系统自动计算投放|'
            r'入网日期晚于(?:\d{4}年)?\d{1,2}月\d{1,2}日客户不参与测算', '', remainder)
        known_tail = not re.sub(r'[，,。；;、：:（）()]', '', remainder)
        whole_price = bool(has_price and global_price and known_tail and not scoped and not uncertain and not rules and not positive)
        choice, category, reason = 'keep', 'normal', '正常导出'
        if whole_stop:
            choice, category, reason = 'omit', 'stop', '整品规停投或无库存，建议不导出'
        elif whole_price:
            choice, category, reason = 'omit', 'price', '整品规按价位段投放，建议不导出'
        elif has_price:
            category = 'review'
            reason = ('已人工填写非零数量，默认保留，请复核' if positive else
                      '含价位段投放表述，可能为局部或混合策略，默认保留，请复核')
        elif product['warnings']:
            category, reason = 'review', '存在校验提示，默认保留，请复核'
        result.append({
            'sourceRow': product['sourceRow'], 'productCode': product['code'],
            'productName': product['name'], 'strategyText': product['strategyText'],
            'originalText': product['originalText'], 'warnings': product['warnings'][:],
            'category': category, 'defaultChoice': choice, 'reason': reason,
        })
    return result
