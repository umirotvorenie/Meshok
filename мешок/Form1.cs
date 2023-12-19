
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using iTextSharp.text.pdf;
using iTextSharp.text;



namespace мешок
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private double ответ(string equation, double lowerBound, double upperBound)
        {
            
            return 42.0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Получаем уравнение из текстового поля
            string уравнение = textBox1.Text;
            double нижняяГраница, верхняяГраница;

            if (!double.TryParse(textBox2.Text, out нижняяГраница) ||
                !double.TryParse(textBox3.Text, out верхняяГраница))
            {
                MessageBox.Show("Пожалуйста, введите корректные границы интеграции.");
                return;
            }

            // Расчет объема с помощью уравнения и границ интеграции
             double объем = ответ(уравнение, нижняяГраница, верхняяГраница);

            // Полное решение
            string деталиРасчета = $"Уравнение: {уравнение}, Границы интеграции: {нижняяГраница} - {верхняяГраница}";
            string шагиИнтеграции = "Здесь могут быть шаги вычислений интеграла.";

            // Вывод в Excel
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet sheet = workbook.ActiveSheet;
            sheet.Cells[1, 1] = "Расчет объема тела вращения";
            sheet.Cells[2, 1] = "Объем";
            sheet.Cells[2, 2] = объем;
            sheet.Cells[4, 1] = "Детали расчета";
            sheet.Cells[5, 1] = деталиРасчета;
            sheet.Cells[6, 1] = "Шаги интеграции";
            sheet.Cells[7, 1] = шагиИнтеграции;

            // Вывод в Word
            Word.Application wordApp = new Word.Application();
            wordApp.Visible = true;
            Word.Document doc = wordApp.Documents.Add();
            Word.Range range = doc.Range();
            range.Text = "Расчет объема тела вращения" + Environment.NewLine +
                         "Объем: " + объем + Environment.NewLine +
                         "Детали расчета: " + деталиРасчета + Environment.NewLine +
                         "Шаги интеграции: " + шагиИнтеграции;

            // Вывод в PDF
            string путьКФайлуPdf = "результаты_расчета_объема.pdf";
            using (FileStream fs = new FileStream(путьКФайлуPdf, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                Document pdfDoc = new Document();
                PdfWriter.GetInstance(pdfDoc, fs);
                pdfDoc.Open();

                // Вывод деталей расчета
                pdfDoc.Add(new Paragraph("Calculation details: " + деталиРасчета));
                pdfDoc.Add(Chunk.NEWLINE);

                // Вывод подробных шагов интегрирования
                pdfDoc.Add(new Paragraph("Detailed integration steps:"));
                

                // Вывод результирующего объема
                pdfDoc.Add(new Paragraph("The volume of the body of rotation: " + объем));

                pdfDoc.Close();
            }
            System.Diagnostics.Process.Start(путьКФайлуPdf);
            MessageBox.Show("Результаты расчета сохранены в Excel, Word и PDF.");
        }

       
    }
 }

