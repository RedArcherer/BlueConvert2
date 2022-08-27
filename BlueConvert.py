import pandas, math
from pathlib import Path
from datetime import datetime
from pytz import timezone
import time

# --- #


class Parser:
    """Class that handles parsing raw CSV Data along with a template into ready to use XLSX files."""

    def __init__(self, data_path: Path, output_path: Path):
        self.output_df = pandas.read_excel("template.xlsx")
        self.data_df = pandas.read_csv(str(data_path.absolute())).dropna(
            subset=["Shipping Name"]
        )

        # if not output_path.is_file():
        #     raise Exception("Parameter output_path must be a folder!")

        self.output_path = output_path

    def parse_phone_number(self, v):
        """Parses output into legible phone number.
        - Strips country code (gets last 10 digits)
        - Converts Mangled Scientific Notation to Int (without ast.literal_eval)
        """
        v = str(v).replace(" ", "").strip()[-10:]
        try:
            int(v)
            return v
        except:
            ...
        if "nan" in v.lower():
            return None
        numbers = v.lower().split("e")
        numbers[1] = numbers[1].replace("+", "")
        return str(int(float(numbers[0]) * (10 ** int(numbers[1])))).lstrip("1")

    def parse(self):
        output_df = self.output_df
        data_df = self.data_df

        output_df["CreditReferenceNo"] = data_df["Name"].apply(lambda v: v[1:])
        output_df["ConsigneeName"] = output_df["ConsigneeAttention"] = data_df[
            "Shipping Name"
        ]
        output_df["ConsigneeAddress1"] = data_df["Shipping Street"]
        output_df["ConsigneePincode"] = data_df["Shipping Zip"].apply(
            lambda v: str(v)[-6:]
        )
        output_df["Consignee Mobile"] = data_df["Shipping Phone"].apply(
            self.parse_phone_number
        )
        output_df["Declared Value"] = data_df["Total"]

        output_df.assign(
            PickupDate=datetime.now(tz=timezone("Asia/Calcutta")).strftime(
                "%d-%m-%Y 12:00:00"
            )
        )

        output_df["ProductCode"] = output_df["ProductCode"].map(
            lambda *args, **kwargs: "D"
        )
        output_df["ProductType"] = output_df["ProductType"].map(
            lambda *args, **kwargs: "NDOX"
        )
        output_df["PieceCount"] = output_df["PieceCount"].map(lambda *args, **kwargs: 1)
        output_df["InvoiceNo"] = output_df["InvoiceNo"].map(lambda *args, **kwargs: 0)
        output_df["PickupTime"] = output_df["PickupTime"].map(
            lambda *args, **kwargs: 1600
        )
        output_df["OriginArea"] = output_df["OriginArea"].map(
            lambda *args, **kwargs: "IMP"
        )
        output_df["CustomerCode"] = output_df["CustomerCode"].map(
            lambda *args, **kwargs: "000206"
        )
        output_df["CustomerName"] = output_df["CustomerName"].map(
            lambda *args, **kwargs: "MADAKE BAMBOO SOLUTIONS LLP"
        )
        output_df["CustomerAddress1"] = output_df["CustomerAddress1"].map(
            lambda *args, **kwargs: "Kasturi Building"
        )
        output_df["CustomerAddress2"] = output_df["CustomerAddress2"].map(
            lambda *args, **kwargs: "Thangal Bazar"
        )
        output_df["CustomerAddress3"] = output_df["CustomerAddress3"].map(
            lambda *args, **kwargs: "Imphal"
        )
        output_df["CustomerTelephone"] = output_df["CustomerTelephone"].map(
            lambda *args, **kwargs: "6374679609"
        )
        output_df["CustomerMobile"] = output_df["CustomerMobile"].map(
            lambda *args, **kwargs: "6374679609"
        )
        output_df["Sender"] = output_df["Sender"].map(
            lambda *args, **kwargs: "MADAKE BAMBOO SOLUTIONS LLP"
        )
        output_df["IsToPayCustomer"] = output_df["IsToPayCustomer"].map(
            lambda *args, **kwargs: "FALSE"
        )

        # Get number of parts/chunks... by dividing the df total number of rows by expected number of rows plus 1
        # source: https://umar-yusuf.blogspot.com/2020/11/split-dataframe-into-chunks.html
        expected_rows = 24
        chunks = math.floor(len(output_df["ConsigneeName"]) / expected_rows + 1)

        i = 0
        j = expected_rows
        for idx, x in enumerate(range(chunks)):
            df_sliced = output_df[i:j]

            df_sliced.to_excel(self.output_path / ("UPLOAD" + str(x+1) + " " + datetime.now(tz=timezone("Asia/Calcutta")).strftime("%d-%m-%Y") + ".xlsx"))

            i += expected_rows
            j += expected_rows
        print("\nCompleted successfully, please check desktop for files")
        time.sleep(60)

if __name__ == "__main__":
    filepath = input("Please drag and drop your orders_export.csv here and then click enter: ")
    filepath = filepath.strip('"')

    p = Parser(
        data_path=Path(filepath),
        output_path=Path.home() / 'Desktop'
    )
    p.parse()