import tkinter as tk
from tkinter import filedialog

import pandas as pd

import matplotlib.pyplot as plt

import os

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.styles import ParagraphStyle


def open_file_dialog():
    file_path = filedialog.askopenfilename(title="Please select the sales file to analyse",
                                           filetypes=[("All files", "*.csv")])
    if file_path:
        print("Selected file:", file_path)

        df = pd.read_csv(file_path)
        required_columns = {'year', 'month', 'sales', 'expenditure'}
        if set(df.columns) == required_columns:
            root.destroy()
            df['profit/loss'] = df['sales'] - df['expenditure']
            year = df['year'].iloc[0]
            df = df.drop(columns=['year'])
            print('Sales data for the year {}:'.format(year))
            print(df.to_string(index=False))
            df['month'] = df['month'].str.title()

            max_sales_month = df['sales'].idxmax()
            row_with_max_sales = df.loc[max_sales_month]
            print(f"The month with the highest sales is {row_with_max_sales['month']} "
                  f"with sales of £{row_with_max_sales['sales']}")
            min_sales_month = df['sales'].idxmin()
            row_with_min_sales = df.loc[min_sales_month]
            print(f"The month with the lowest sales is {row_with_min_sales['month']} "
                  f"with sales of £{row_with_min_sales['sales']}")
            average_sales = df['sales'].mean()
            average_sales = round(average_sales, 2)
            print(f"The average sales is: £{average_sales}")
            total_sales = df['sales'].sum()
            print(f"The annual sales is: £{total_sales}")

            sales_expenditure_graph_path = 'sales_expenditure_graph.png'
            profit_graph_path = 'profit_graph.png'

            def plot_sales_expenditure_graph(tbl):
                plt.figure(figsize=(10, 6))
                plt.bar(tbl['month'], tbl['sales'], label='Sales', color='blue', alpha=0.7)
                plt.bar(tbl['month'], tbl['expenditure'], label='Expenditure', color='red', alpha=0.7)
                plt.title('Sales and Expenditure - {}'.format(year))
                plt.xlabel('Month')
                plt.ylabel('Amount (£)')
                plt.legend()
                plt.grid(True)
                plt.savefig(sales_expenditure_graph_path)
                plt.close()

            def plot_profit_graph(tbl):
                plt.figure(figsize=(10, 6))
                plt.bar(tbl['month'], tbl['profit/loss'], label='Sales', color='blue', alpha=0.7)
                plt.title('Monthly Profits & Losses - {}'.format(year))
                plt.xlabel('Month')
                plt.ylabel('Profit/Loss (£)')
                plt.grid(True)
                plt.savefig(profit_graph_path)
                plt.close()

            plot_sales_expenditure_graph(df)
            plot_profit_graph(df)

            column_name_mapping = {'month': 'Month',
                                   'sales': 'Sales (£)',
                                   'expenditure': 'Expenditure (£)',
                                   'profit/loss': 'Profit/Loss (£)',
                                   }
            df.rename(columns=column_name_mapping, inplace=True)

            pdf_filename = 'sales_report_{}.pdf'.format(year)
            if os.path.exists(pdf_filename):
                try:
                    os.remove(pdf_filename)
                except PermissionError:
                    print(f"Error: Please close the PDF file '{pdf_filename}' before running this program.")
                    exit()

            doc = SimpleDocTemplate(pdf_filename, pagesize=letter)

            title_style = getSampleStyleSheet()['Title']
            title_text = 'Financial Report - {}'.format(year)

            data = [df.columns.tolist()] + df.values.tolist()
            table = Table(data)
            style = TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                                ('GRID', (0, 0), (-1, -1), 1, colors.black)])
            table.setStyle(style)

            summary_text = f"""
                    <br />
                    This report presents the financial data for the year {year}. The table provides an overview of 
                    sales, expenditure, and profits and losses for each month, while the graphs below illustrate trends 
                    of these key figures over the year. This helps in understanding and analysing business performance.
                    <br /><br />
                    \u2022 The month with the highest sales is {row_with_max_sales['month']} with sales of 
                    £{row_with_max_sales['sales']}.<br />
                    \u2022 The month with the lowest sales is {row_with_min_sales['month']} with sales of 
                    £{row_with_min_sales['sales']}.<br />
                    \u2022 The average sales is: £{average_sales}.<br />
                    \u2022 The annual sales is: £{total_sales}.<br />
                    """
            summary_style = ParagraphStyle(
                name='BulletStyle'
            )

            content = [Paragraph(title_text, title_style), Spacer(1, 12), table, Paragraph(summary_text, summary_style),
                       Image(sales_expenditure_graph_path, width=400, height=300),
                       Image(profit_graph_path, width=400, height=300)]

            doc.build(content)

            print(f"PDF report '{pdf_filename}' has been generated.")
        else:
            print("The table does not have the required columns: 'year', 'month', 'sales', 'expenditure'. "
                  "Please select a file containing this data.")


root = tk.Tk()
root.title("File Selector")
root.geometry("300x100")

button = tk.Button(root, text="Select File", command=open_file_dialog)
button.pack(pady=20)

root.mainloop()
