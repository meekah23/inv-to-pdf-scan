import sys
import pandas as pd
import sqlite3
import time
from itertools import groupby
from collections import defaultdict
from PyQt5.QtWidgets import *
from PyQt5 import uic
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import BaseDocTemplate, Table, TableStyle, Paragraph, Frame, PageTemplate, PageBreak, SimpleDocTemplate
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont



pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
pdfmetrics.registerFont(TTFont('ArialBd', 'ArialBd.ttf'))





class VanInvoice(QWidget):
    def __init__(self):
        super(VanInvoice, self).__init__()
        uic.loadUi("UI/vaninv.ui",self)
        self.show()
        self.scan.clicked.connect(self.getData)
        self.pdfbtn.clicked.connect(self.toPDF)
        """
        #load excel file
        conn = sqlite3.connect(':memory:')
        cur = conn.cursor()
        df = pd.read_excel("Data/PR_HEAD.xlsx",sheet_name="PR_HEAD")
        df.to_sql(name='PR_HEAD', con=conn, if_exists='append', index=False)
        cur.execute("SELECT OINO FROM PR_HEAD")
        inv = cur.fetchall()
        list1 = [', '.join(map(str, x)) for x in inv]
        completer = QCompleter(list1)
        self.lineEdit.setCompleter(completer)
        """





    def getData(self):
        try:
            start_time = time.time()
            global table, headers,result,right
            conn = sqlite3.connect(':memory:')
            cur = conn.cursor()

            # load excel file
            df = pd.read_excel("Data/PR_HEAD.xlsx", sheet_name="PR_HEAD")
            df.to_sql(name='PR_HEAD', con=conn, if_exists='append')
            df1 = pd.read_excel("Data/PR_DETL.xlsx", sheet_name="PR_DETL")
            df1.to_sql(name='PR_DETL', con=conn, if_exists='append')

            #get headers
            cur.execute("SELECT OIDATE,OINO,OSLSMAN,ADDEDBY,OQNO FROM PR_HEAD WHERE PR_STATUS is null",)
            result = cur.fetchall()
            #headers = [[x for x in g] for x, g in groupby(result, key = lambda x: x[3])]
            #print(headers)
            cur.execute("SELECT OINO,OQNO FROM PR_HEAD WHERE PR_STATUS is null",)
            right = cur.fetchall()
            
            #get inv details
            cur.execute("SELECT OD_ITEM, "
                        "CASE WHEN round(OD_QTY,2) = 0 AND round(OD_PRICE,2) = 0 AND round(OD_AMOUNT,2) = 0 THEN ''"
                        "ELSE round(OD_QTY)"
                        "END, "
                        "COALESCE(OD_DESCR, ' ') as descr,"
                        "CASE WHEN round(OD_QTY,2) = 0 AND round(OD_PRICE,2) = 0 AND round(OD_AMOUNT,2) = 0 THEN '' "
                        "ELSE round(OD_PRICE,2)"
                        "END, "
                        "CASE WHEN round(OD_QTY,2) = 0 AND round(OD_PRICE,2) = 0 AND round(OD_AMOUNT,2) = 0 THEN '' "
                        "ELSE round(OD_AMOUNT,2) "
                        "END, "
                        "PR_HEAD.OINO "
                        #"sum(round(OD_AMOUNT,2)), sum(round(PR_HEAD.OGST,2)), sum(round(OD_AMOUNT,2))+ sum(round(PR_HEAD.OGST,2))"
                       "FROM PR_DETL "
                        "INNER JOIN PR_HEAD ON PR_DETL.OD_UNO = PR_HEAD.ONUMBER "
                        "WHERE PR_HEAD.PR_STATUS is null")
            table = cur.fetchall()
            #values = [[k,[x[1:] for x in g]] for x, g in groupby(table, key = lambda x: x[5])]
            #values = [[x for x in g] for x, g in groupby(table, key = lambda x: x[5])]
            #dflist = pd.DataFrame(table, columns =["Code", "Qty", "Description", "Unit Price", "Amount", "Invoice"])
            #dflist['match'] = dflist.Invoice.eq(dflist.Invoice.shift()) 
            #sep = dict(list(dataframe.groupby("Invoice")))
            #print(values)
            #print(sep[358826])

            self.tableWidget.setRowCount(0)
            for row_number, row_data in enumerate(table):
                self.tableWidget.insertRow(row_number)
                for column_number, data in enumerate(row_data):
                    self.tableWidget.setItem(row_number, column_number, QTableWidgetItem(str(data)))

            #get total
            

            end_time = time.time()
            execution_time = end_time - start_time
            self.time_label.setText("Data fetched for " + str(execution_time) + " secs")
            

            QMessageBox.information(self, "Done!", "Data fetched.")

        except Exception as e:
                print(e)
                QMessageBox.information(self,"Error", "Failed to run script.")

    def toPDF(self):
        start_time = time.time()
        filename, _ = QFileDialog.getSaveFileName(self, 'Save file', '', 'PDF File (*.pdf)')
        if filename != '':
            try:

                style2 = ParagraphStyle(
                    name='Normal',
                    fontName='ArialBd',
                    fontSize=22,
                    alignment=1,
                    spaceAfter=15,
                )
                style3 = ParagraphStyle(
                    name='Normal',
                    fontName='Arial',
                    fontSize=12,
                    alignment=1,
                    spaceBefore=20,
                )

                addstyle = ParagraphStyle(
                    name='Normal',
                    fontName='ArialBd',
                    fontSize=10,
                    alignment=1,
                )

                tablestyle1 = TableStyle([
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                    ('FONTNAME', (0, 0), (-1, -1), 'ArialBd'),
                    #('BOX', (0, 0), (-1, -1), 0.5, colors.black),

                 ])

                tablestyle2 = TableStyle([

                     ('FONTSIZE', (0, 0), (-1, -1), 10),
                     ('FONTNAME', (0, 0), (-1, -1), 'ArialBd'),
                     ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
                     ('LINEBEFORE', (1, 0), (1, -1), 0.5, colors.black)

                 ])

                headerstyle = TableStyle([

                     ('FONTSIZE', (0, 0), (-1, 0), 9),
                     ('FONTNAME', (0, 0), (-1, 0), 'ArialBd'),
                     ('LINEABOVE', (0, 0), (-1, -1), 1, colors.black),
                     ('LINEBEFORE', (0, 0), (0, -1), 1, colors.black),
                     ('LINEAFTER', (4, 0), (4, -1), 1, colors.black),
                     ])

                footerstyle = TableStyle([

                     ('FONTSIZE', (0, 0), (-1, -1), 8),
                     ('FONTSIZE', (1, 0), (2, -1), 9),
                     ('FONTNAME', (0, 0), (-1, -1), 'ArialBd'),
                     ('LINEBEFORE', (0, 0), (0, 4), 1, colors.black),
                     ('LINEBEFORE', (2, 0), (-1, -1), 1, colors.black),
                     ('LINEAFTER', (2, 0), (-1, -1), 1, colors.black),
                     ('LINEABOVE', (0, 5), (-1, 5), 1, colors.black),
                     ('LINEBELOW', (2, 5), (-1, 5), 1, colors.black),
                     ('ALIGN', (2, 0), (2, -1), "RIGHT"),
                      ])

                def header(canvas, pdf):
                   
                        # Draw heading
                        heading = Paragraph("VANCOUVER GLASS (1990) LTD.", style2)
                        heading.wrap(pdf.width, inch * 0.3)
                        heading.drawOn(canvas, pdf.leftMargin, pdf.height + inch)

                        # Draw subheading.
                        subheading = Paragraph("INVOICE", style3)
                        subheading.wrap(pdf.width, inch * 0.2)
                        subheading.drawOn(canvas, pdf.leftMargin, pdf.height + inch * 0.5)
                    
                        #table line
                        canvas.line(1.7 * inch, 2.5 *inch, 1.7 *inch, 7.6 *inch)
                        canvas.line(2.26 * inch, 2.5 * inch, 2.26 * inch, 7.6 * inch)
                        canvas.line(6.41 * inch, 2.5 * inch, 6.41 * inch, 7.6 * inch)
                        canvas.line(7.19 * inch, 2.5 * inch, 7.19 * inch, 7.6 * inch)

                        tablelist1 = [["Date:","Invoice#: "],
                                  ["Salesman:"],
                                  ["GST#: 121989834RT","Quote#: "],
                                  ["Added-by:","Page: "+"%d " % doc.page]
                                  ] 
                        """
                        tabledetails = [] 
                        for row in range(len(values)):
                            for col in range(len(values[0])):
                                table_item = values[row][col]              
                                tabledetails.append('' if table_item is None else str(table_item))

                        print(tabledetails)"""
                        
                 
                        tablelist2 = [["Sold To:", "Location: "],
                                  ["", "Ship To:"],
                                  ["", ""],
                                  ["", ""],
                                  ["", ""],
                                  ["", ""],
                                  ["Phone: ", ""],
                                  ["PST Exempt# "]]

                        headerlist = [["Code", "Qty", "Description", "Unit Price", "Amount"]]

                        footerlist = [["", "Sub-Total:", ""],
                                 ["Overdue accounts will be charged 2% per month.", "GST", ""],
                                 ["Please enclose a copy of the invoice with the cheque."],
                                 [""],
                                 [""],
                                 ["Charge to Acount", "Total Amount:", ""]]

                        table1 = Table(tablelist1, colWidths=[400,200], rowHeights=[10, 10, 10, 10],
                                   hAlign='CENTER', spaceBefore=5, style=tablestyle1)
                        table1.wrap(pdf.width, inch)
                        table1.drawOn(canvas, pdf.leftMargin, pdf.height - inch * 0.4)  
                        
                        table2 = Table(tablelist2, colWidths=[268, 268], rowHeights=[20, 15, 10, 10, 10, 10, 10, 20],
                                   hAlign='CENTER', spaceBefore=5, style=tablestyle2)
                        table2.wrap(pdf.width, inch)
                        table2.drawOn(canvas, pdf.leftMargin, pdf.height - inch * 2)

                        header = Table(headerlist,hAlign='LEFT', spaceBefore=5, repeatRows=1, style=headerstyle,
                                     colWidths=[86, 43, 301, 54, 53])
                        header.wrap(pdf.width, inch)
                        header.drawOn(canvas, pdf.leftMargin, pdf.height - inch * 2.35)
                        footertable = Table(footerlist, colWidths=[410, 71, 56], rowHeights=[20, 10, 10, 10, 10, 20],
                                        hAlign='LEFT', spaceBefore=5, style=footerstyle)
                        footertable.wrap(pdf.width, inch)
                        footertable.drawOn(canvas, pdf.leftMargin, 1.38 * inch)


                        addressnote = Paragraph("1706 E. HASTINGS, VAN, B.C. V5L 1S9 Phone (604)253-7707 Fax (604)253-8448",addstyle)
                        addressnote.wrap(pdf.width, inch)
                        addressnote.drawOn(canvas, pdf.leftMargin, 0.7 * inch)
                        

                doc = BaseDocTemplate(filename, leftMargin=0.5 * inch, rightMargin=0.5 * inch)
                
                frame = Frame(

                    0.5 * inch,  # x
                    2.5 * inch,  # y at bottom
                    7.46 * inch,  # width
                    4.83 * inch,  # height
                    showBoundary=1
                )
                
                #template = PageTemplate(id='all_pages', frames=frame, onPage=header)

                template = PageTemplate(id='all_pages',frames=frame,onPage=header)
                doc.addPageTemplates([template])

                tablestyle3 = TableStyle([
                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('FONTNAME', (0, 1), (-1, -1), 'Arial'),
                    ('ALIGN',(1,0),(1,-1), "RIGHT"),
                    ('ALIGN', (3, 0), (5, -1), "RIGHT"),
                ])
                table_style = TableStyle([
                    #('BACKGROUND', (1,1), (-2,-2), colors.green),
                    #('TEXTCOLOR', (0,0), (1,-1), colors.red),
                    ('BOX', (0,0), (-1,-1), 0.45, colors.black),
                    ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.blue),
                ])

                all_data = [[x for x in g] for x, g in groupby(table, key = lambda x: x[5])]
                Elements = []

                for data in all_data:
                    table3 = Table(data, style=tablestyle3, hAlign='LEFT',repeatRows=1,colWidths=[80, 40, 300, 55, 50, 100], rowHeights=10)
                    #table.setStyle(table_style)
                    Elements.append(table3)
                    Elements.append(PageBreak())


                               
                doc.build(Elements)

                end_time = time.time()
                execution_time = end_time - start_time
                self.time_label.setText("PDF created for " + str(execution_time) + " secs")
                QMessageBox.information(self, "Done!", "File Exported.")

            except Exception as e:
                print(e)
                QMessageBox.information(self, "Error", "Failed to run script.")


app = QApplication(sys.argv)
design = VanInvoice()

app.exec_()


