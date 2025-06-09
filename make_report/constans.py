patterns = {
    "price_from_sku": r'[-рсолхэпик ]\s*(\d+)',
    "payment_option_name": r"Пром-оплата|олхрс|оплрс|ОплР\.С",
    "denis_positive": r'денис-(\d+)',
    "denis_negative": r'взялиидениса|денис(\d+)',
    "marad": r"дропмард-(\d+)",
    "barcode": r'[0-9]{12,14}',
    "pe": r'пе(\d+)'
}
