from io import BytesIO
from pathlib import Path
import numpy as np
import matplotlib.pyplot as plt

from viktor import ViktorController
from viktor.parametrization import ViktorParametrization, TextField, DateField, Text, LineBreak, Table, NumberField, DownloadButton
from viktor.views import PDFView, PDFResult
from viktor.external.word import render_word_file, WordFileTag, WordFileImage
from viktor.utils import convert_word_to_pdf
from viktor.result import DownloadResult


class Parametrization(ViktorParametrization):

    intro = Text("# Invoice app 💰 \n This app makes an invoice based on your own Word template")

    client_name = TextField("Name of client")
    company = TextField("Company name")
    lb1 = LineBreak()  # This is just to separate fields in the parametrization UI
    date = DateField("Date")

    # Table
    table_price = Table("Products")
    table_price.qty = NumberField("Quantity", suffix="-", min=0)
    table_price.desc = TextField("Description", suffix="-")
    table_price.price = NumberField("Price", suffix="€", min=0)

    # Downloadbutton
    download_word_file = DownloadButton('Download report', method='download_word_file')


class Controller(ViktorController):

    label = 'reporting'
    parametrization = Parametrization

    def generate_word_document(self, params):

        total_price = self.calc_total_price(params)
        table = self.process_table(params)

        # Create emtpy components list to be filled later
        components = []

        # Fill components list with data
        components.append(WordFileTag("Client_name", params.client_name))
        components.append(WordFileTag("company", params.company))
        components.append(WordFileTag("date", str(params.date))) # Convert date to string format

        components.append(WordFileTag("total_price", str(total_price)))  # Convert price float to string
        components.append(WordFileTag("table1", table))
        components.append(WordFileTag("table2", table))

        # Place image
        figure = self.create_figure(params)
        word_file_figure = WordFileImage(figure, "figure_sales", width=500)
        components.append(word_file_figure)

        # Get path to template and render word file
        template_path = Path(__file__).parent / "files" / "Template.docx"
        with open(template_path, 'rb') as template:
            word_file = render_word_file(template, components)

        return word_file

    @PDFView("PDF viewer", duration_guess=5)
    def pdf_view(self, params, **kwargs):
        word_file = self.generate_word_document(params)

        with word_file.open_binary() as f1:
            pdf_file = convert_word_to_pdf(f1)

        return PDFResult(file=pdf_file)

    @staticmethod
    def calc_total_price(params):
        # Get user entry from params
        product_table = params.table_price

        # Calculate total price from quantities and unit price
        quantities = [row["qty"] for row in product_table]
        prices = [row["price"] for row in product_table]
        total_price = 0
        for qty, price in zip(quantities, prices):
            total_price += qty * price

        return total_price

    def process_table(self, params):
        total_price = self.calc_total_price(params)
        product_table = params.table_price
        for row in product_table:
            row["total"] = row["qty"] * row["price"]
            row["perc"] = str(round((row["total"] / total_price) * 100, 2)) + "%"

        return product_table

    @staticmethod
    def create_figure(params):
        product_table = params.table_price
        # Create figure
        fig, ax = plt.subplots(figsize=(16, 8))
        products = [row["desc"] for row in product_table]
        qty = [np.round(row["qty"], 2) for row in product_table]
        ax.pie(qty, labels=products, autopct="%1.1f%%")
        ax.set_title("Pie chart total sold products")
        png_data = BytesIO()
        fig.savefig(png_data, format='png')
        plt.close()

        return png_data

    def download_word_file(self, params, **kwargs):
        word_file = self.generate_word_document(params)

        return DownloadResult(word_file, "Invoice.docx")