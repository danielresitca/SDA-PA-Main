import xml.etree.ElementTree as ET

def extract_lines(xml_path):
    tree = ET.parse(xml_path)
    root = tree.getroot()
    ns = {'cbc': 'urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2',
          'cac': 'urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2'}
    lines = []
    for i, item in enumerate(root.findall('.//cac:InvoiceLine', ns), 1):
        desc = item.findtext('cac:Item/cbc:Description', default='', namespaces=ns)
        qty = item.findtext('cbc:InvoicedQuantity', default='0', namespaces=ns)
        price = item.findtext('cac:Price/cbc:PriceAmount', default='0', namespaces=ns)
        total = item.findtext('cbc:LineExtensionAmount', default='0', namespaces=ns)
        lines.append({
            "line_id": str(i),
            "description": desc,
            "quantity": qty,
            "unit_price": price,
            "line_total": total
        })
    return lines