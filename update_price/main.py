import datetime
import json
import logging
import math
import os
import re
import pprint
import colorlog
import ezsheets
from dotenv import load_dotenv
from openpyxl import load_workbook

from paths import prom_grand, roz_grand, epic_grand

load_dotenv()
handler = colorlog.StreamHandler()
handler.setFormatter(
    colorlog.ColoredFormatter(
        "%(log_color)s%(levelname)s:%(name)s:%(message)s",
        log_colors={
            "DEBUG": "cyan",
            "INFO": "green",
            "WARNING": "yellow",
            "ERROR": "red",
            "CRITICAL": "bold_red",
        },
    )
)

logger = colorlog.getLogger("my_logger")
logger.addHandler(handler)
logger.setLevel(logging.DEBUG)

logging.getLogger().handlers.clear()


class UpdatePrice:
    def __init__(
        self,
        export_file_path="",
        margin=100,
        original_margin=400,
        currency_rate=40,
        rate_sell=15,
        vendor_code_column="D",
        valuta="USD",
        rrp_column="G",
    ):
        self.export_file_path = export_file_path
        self.margin = margin
        self.original_margin = original_margin
        self.currency_rate = currency_rate
        self.rate_sell = rate_sell
        self.vendor_code_column = vendor_code_column
        self.valuta = valuta
        self.rrp_column = rrp_column

    @staticmethod
    def royalty(price, rate):
        return math.ceil((price / (100 - rate)) * 100)

    def __vendor_code(self, worksheet, row_number, rate):
        vendor_code = worksheet[f"{self.vendor_code_column}{row_number}"].value

        if vendor_code is None:
            return None, None

        is_original = vendor_code.startswith("OR|")
        has_double_pipe = "||" in vendor_code
        has_T = "Т" in vendor_code

        separator = "||" if has_double_pipe else "|"

        try:
            pp = float(vendor_code.split(separator)[-1])
        except (ValueError, IndexError):
            logger.error(
                f"Error parsing vendor code at row {row_number}: {vendor_code}"
            )
            return None, None

        base_price = pp * self.currency_rate if has_double_pipe else pp
        margin = self.original_margin if is_original else self.margin
        base_price += margin + (150 if has_T else 0)

        new_price = self.royalty(base_price, rate)
        old_price = self.royalty(new_price, self.rate_sell)

        return new_price, old_price

    def __availability(self, worksheet, row_number):

        availability = worksheet[f"P{row_number}"].value

        if availability in ("+", "!"):
            discount = f"{self.rate_sell}%"
            date_start = datetime.datetime.now().strftime("%d.%m.%Y")
            date_end = (datetime.datetime.now() + datetime.timedelta(days=7)).strftime(
                "%d.%m.%Y"
            )
        else:
            discount = ""
            date_start = ""
            date_end = ""

        return discount, date_start, date_end

    @staticmethod
    def __get_rate(export_rate_id, rate_file="prom_rate.json"):
        try:
            with open(rate_file, "r", encoding="utf-8") as f:
                rates = json.load(f)

            for item in rates:
                if item.get("cat_id") == str(export_rate_id):
                    rate_str = item.get("rate", "0%")
                    return float(rate_str.rstrip("%"))
            return 0.0
        except Exception as ex:
            logger.error(f"Error reading rate from {rate_file}: {ex}")
            return 0.0

    @staticmethod
    def __vendor_validation(vc_export):

        if vc_export is None:
            return False

        if re.match(r"^Код_товара", vc_export):
            return False

        return True

    def __get_price(self, sheet_id, vendor_code_export):
        try:
            ss = ezsheets.Spreadsheet(sheet_id)
            sheet = ss[0]
            data = sheet.getRows()

            vendor_code_dict = {row[1]: row for row in data}

            if "OR|" in vendor_code_export:
                vendor_code_key = vendor_code_export.split("|")[1]
            else:
                vendor_code_key = vendor_code_export.split("|")[0]

            if vendor_code_key in vendor_code_dict:
                row = vendor_code_dict[vendor_code_key]
                price = row[4] if self.valuta == "USD" else row[5]
                availability = row[8]

                price = price.replace("$", "").replace(",", ".").strip()
                return float(price), availability, vendor_code_key
            else:
                return None
        except Exception as e:
            logger.error(f"Error in __get_price: {e}")
            return None

    @staticmethod
    def __get_rrp(sheet_id, vendor_code_export):
        try:
            ss = ezsheets.Spreadsheet(sheet_id)
            sheet = ss[0]
            data = sheet.getRows()

            vendor_code_dict = {row[1]: row for row in data}

            if "OR|" in vendor_code_export:
                vendor_code_key = vendor_code_export.split("|")[1]
            else:
                vendor_code_key = vendor_code_export.split("|")[0]

            if vendor_code_key in vendor_code_dict:
                row = vendor_code_dict[vendor_code_key]
                rrp = row[6]
                return float(rrp)
            else:
                logger.warning(
                    f"Vendor code {vendor_code_export} not found in RRP sheet."
                )
                return None
        except Exception as e:
            logger.error(f"Error in __get_rrp: {e}")
            return None

    def __compare_and_set_price(self, worksheet, row_number, sheet_id, final_price):
        rrp = self.__get_rrp(
            sheet_id, worksheet[f"{self.vendor_code_column}{row_number}"].value
        )
        if rrp is not None and final_price < rrp:
            final_price = rrp
        return final_price

    def __put_id_prom(self, worksheet, row_number, sheet_id):
        vendor_code_export = worksheet[f"{self.vendor_code_column}{row_number}"].value

        result = self.__get_price(
            sheet_id=sheet_id, vendor_code_export=vendor_code_export
        )

        if result:
            price, availability, vendor_code = result

            worksheet[f"P{row_number}"].value = "!" if availability == "TRUE" else "-"

            separator = "||" if "||" in vendor_code_export else "|"

            prefix = "OR|" if vendor_code_export.startswith("OR|") else ""
            worksheet[f"{self.vendor_code_column}{row_number}"].value = (
                f"{prefix}{vendor_code}{separator}00{price}"
            )
            return True
        else:
            return False

    def updateProm(self, rate_column="AA", from_price=None):
        wb = load_workbook(filename=self.export_file_path)
        ws = wb.active

        for i in range(2, ws.max_row + 1):
            vendor_code_export = ws[f"{self.vendor_code_column}{i}"].value

            if not self.__vendor_validation(vendor_code_export):
                continue

            export_rate_id = ws[f"{rate_column}{i}"].value
            rate = self.__get_rate(export_rate_id)

            if from_price and not self.__put_id_prom(ws, i, from_price):
                logger.warning(f"Product at row {i} / {ws.max_row} was not found.")

            new_price, old_price = self.__vendor_code(ws, i, rate)

            if new_price is None:
                continue

            final_price = self.__compare_and_set_price(ws, i, from_price, new_price)

            old_price = self.royalty(final_price, self.rate_sell)

            discount, date_start, date_end = self.__availability(ws, i)

            ws[f"I{i}"].value = str(old_price)
            ws[f"AE{i}"].value = discount
            ws[f"AI{i}"].value = date_start
            ws[f"AJ{i}"].value = date_end

            logger.info(f"Row {i} / {ws.max_row} updated.")

        wb.save(self.export_file_path)

    def updateEpik(self, rate, from_price=None):
        wb = load_workbook(filename=self.export_file_path)
        ws = wb.active

        for i in range(2, ws.max_row + 1):

            if from_price is not None:
                vencod_export = ws[f"{self.vendor_code_column}{i}"].value

                result = self.__get_price(
                    sheet_id=from_price, vendor_code_export=vencod_export
                )

                if result is None:
                    logger.warning(f"Row {i} / {ws.max_row} completed")
                    logger.warning(f"{vencod_export} was not found")
                    continue

                price, availability, vendor_code = result

                if availability == "TRUE":
                    ws[f"H{i}"].value = "в наявності"
                elif availability == "FALSE":
                    ws[f"H{i}"].value = "немає в наявності"

                if re.search(r"\|\|", vencod_export):

                    if re.match(r"^OR\|+", vencod_export):
                        new_price = self.royalty(
                            price * self.currency_rate + self.original_margin, rate
                        )
                        old_price = self.royalty(new_price, self.rate_sell)
                    else:
                        new_price = self.royalty(
                            price * self.currency_rate + self.margin, rate
                        )
                        old_price = self.royalty(new_price, self.rate_sell)

                else:

                    if re.match(r"^OR\|+", vencod_export):
                        new_price = self.royalty(price + self.original_margin, rate)
                        old_price = self.royalty(new_price, self.rate_sell)
                    else:
                        new_price = self.royalty(price + self.margin, rate)
                        old_price = self.royalty(new_price, self.rate_sell)

                ws[f"E{i}"].value = str(new_price)
                ws[f"F{i}"].value = str(old_price)
                logger.info(f"Row {i} / {ws.max_row} completed")

            else:
                new_price, old_price = self.__vendor_code(
                    worksheet=ws, row_number=i, rate=rate
                )

                ws[f"E{i}"].value = str(new_price)
                ws[f"F{i}"].value = str(old_price)
                logger.info(f"Row {i} / {ws.max_row} completed")

        wb.save(f"{self.export_file_path}")

    def updateRozetka(self, rate, from_price=None):
        wb = load_workbook(filename=self.export_file_path)
        ws = wb.active

        for i in range(2, ws.max_row + 1):

            if from_price is not None:
                vencod_export = ws[f"{self.vendor_code_column}{i}"].value

                result = self.__get_price(
                    sheet_id=from_price, vendor_code_export=vencod_export
                )

                if result is None:
                    logger.warning(f"Row {i} / {ws.max_row} completed")
                    logger.warning(f"Vendor code {vencod_export} not found.")
                    continue

                price, availability, vendor_code = result

                if availability == "TRUE":
                    ws[f"P{i}"].value = "Есть в наличии"
                elif availability == "FALSE":
                    ws[f"P{i}"].value = "Нет в наличии"

                if re.search(r"\|\|", vencod_export):

                    if re.match(r"^OR\|+", vencod_export):
                        new_price = self.royalty(
                            price * self.currency_rate + self.original_margin, rate
                        )
                        old_price = self.royalty(new_price, self.rate_sell)
                    else:
                        new_price = self.royalty(
                            price * self.currency_rate + self.margin, rate
                        )
                        old_price = self.royalty(new_price, self.rate_sell)

                    ws[f"I{i}"].value = str(new_price)
                    ws[f"J{i}"].value = str(old_price)

                else:
                    if re.match(r"^OR\|+", vencod_export):
                        new_price = self.royalty(price + self.original_margin, rate)
                        old_price = self.royalty(new_price, self.rate_sell)
                    else:
                        new_price = self.royalty(price + self.margin, rate)
                        old_price = self.royalty(new_price, self.rate_sell)

                    ws[f"I{i}"].value = str(new_price)
                    ws[f"J{i}"].value = str(old_price)

                logger.info(f"Row {i} / {ws.max_row} completed")

            else:
                new_price, old_price = self.__vendor_code(
                    worksheet=ws, row_number=i, rate=rate
                )

                ws[f"I{i}"].value = new_price
                ws[f"J{i}"].value = old_price

                logger.info(f"Row {i} / {ws.max_row} completed")

        wb.save(self.export_file_path)


def prom(
    export_file_path,
    prices,
    margin: int = 70,
    original_margin: int = 300,
    current_course: int | float = 39,
    rate_sell: int = 20,
    valuta: str = "USD",
    vendor_code_column: str = "A",
):
    price_lists = prices

    export = UpdatePrice(
        export_file_path=export_file_path,
        margin=margin,
        original_margin=original_margin,
        currency_rate=current_course,
        rate_sell=rate_sell,
        vendor_code_column=vendor_code_column,
        valuta=valuta,
    )

    for price in price_lists:
        export.updateProm(from_price=os.getenv(price))


def epicentr(
    base_dir,
    prices,
    margin: int = 100,
    original_margin: int = 300,
    current_course: int | float = 39,
    rate_sell: int = 20,
    valuta: str = "USD",
    vendor_code_column: str = "D",
):
    BASE_DIR = base_dir
    PRICE_LISTS = prices

    for path in os.listdir(BASE_DIR):

        path_to_file = os.path.join(BASE_DIR, path)

        export = UpdatePrice(
            export_file_path=path_to_file,
            margin=margin,
            original_margin=original_margin,
            currency_rate=current_course,
            rate_sell=rate_sell,
            vendor_code_column=vendor_code_column,
            valuta=valuta,
        )

        with open("epik_rate.json", "r", encoding="utf-8") as f:
            file = json.load(f)

        for item in file.get("rate"):
            if re.match(rf"{item}", path.split(".")[0]):
                logger.info(f"{item}:{file.get('rate').get(item)}")
                rate = file.get("rate").get(item)
                break

        for price in PRICE_LISTS:
            export.updateEpik(rate=rate, from_price=os.getenv(price))


def rozetka(
    base_dir,
    prices,
    margin: int = 100,
    original_margin: int = 300,
    current_course: int | float = 39,
    rate_sell: int = 20,
    valuta: str = "USD",
    vendor_code_column: str = "E",
):
    BASE_DIR = base_dir
    PRICE_LISTS = prices

    for path in os.listdir(BASE_DIR):

        path_to_file = os.path.join(BASE_DIR, path)

        export = UpdatePrice(
            export_file_path=path_to_file,
            margin=margin,
            original_margin=original_margin,
            currency_rate=current_course,
            rate_sell=rate_sell,
            vendor_code_column=vendor_code_column,
            valuta=valuta,
        )

        with open("rozetka_rate.json", "r", encoding="utf-8") as f:
            file = json.load(f)

        for item in file.get("rate"):

            if re.match(rf"{item}", path.split(".")[0]):
                rate = file.get("rate").get(item)
                logger.info(f"{item}:{file.get('rate').get(item)}")
                logger.info("=" * 100)
                break

        if rate is None:
            rate = 1.0

        for price in PRICE_LISTS:
            (export.updateRozetka(rate=rate, from_price=os.getenv(price)))


def manual(margin, original_margin, rate, rate_sell, currency_rate):
    export = UpdatePrice(currency_rate=currency_rate)
    while True:
        price = float(input("Enter price: "))
        new_price = export.royalty(price + margin, rate)
        old_price = export.royalty(new_price, rate_sell)
        or_new_price = export.royalty(price + original_margin, rate)
        or_old_price = export.royalty(or_new_price, rate_sell)


def main(marketplace):
    match marketplace:
        case "MANUAL":
            manual(
                margin=150,
                original_margin=450,
                rate=15.15,
                rate_sell=20,
                currency_rate=41,
            )

        case "PROM":
            prom(
                export_file_path=prom_grand,
                prices=["GRAND_ELTOS"],
                valuta="USD",
                current_course=43.20,
                margin=100,
                original_margin=430,
                rate_sell=45,
            )
            # prom_press, prom_korm, prom_incubator, prom_gaz, prom_grand

        case "EPICENTR":
            epicentr(
                base_dir=epic_grand,
                prices=["GRAND_ELTOS"],
                valuta="USD",
                current_course=43.20,
                margin=150,
                original_margin=480,
            )
            # epic_press, epic_incubator, epic_korm, epic_gaz, epic_house_tech, epic_grand

        case "ROZETKA":
            rozetka(
                base_dir=roz_grand,
                margin=150,
                original_margin=480,
                prices=["GRAND_ELTOS"],
                valuta="USD",
                current_course=43.20,
                rate_sell=10,
            )
            # roz_press, roz_gaz, roz_incubator,roz_grand

        case _:
            logger.info("You do not have any access to the code")


if __name__ == "__main__":
    # PRESS, KORM, GAZ, INCUBATOR, KITCHEN, GRAND_ELTOS

    MARKETPLACES = ["ROZETKA", "PROM", "EPICENTR"]

    for MARKETPLACE in MARKETPLACES:
        logger.info("=" * 50)
        logger.info(f'{"=" * 20}    {MARKETPLACE}    {"=" * 20}')
        logger.info("=" * 50)
        main(MARKETPLACE)
