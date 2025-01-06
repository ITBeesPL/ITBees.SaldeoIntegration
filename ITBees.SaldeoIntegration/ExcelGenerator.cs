using ITBees.FAS.Payments.Controllers.Operator;
using ITBees.FAS.Payments.Interfaces.Models;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ITBees.SaldeoIntegration;

public class ExcelGenerator
{
    ExcelGenerator()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }
    
    public static void InsertDataIntoTemplate(
        List<FinishedPaymentVm> payments,
        string templateFilePath,
        string outputFilePath, string invoiceSufffix)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        // Load the existing Excel file from disk
        FileInfo templateFile = new FileInfo(templateFilePath);

        // IMPORTANT: License context for EPPlus, if needed:
        // ExcelPackage.LicenseContext = LicenseContext.NonCommercial; 
        // (uncomment if you're not using an official commercial license)

        using (var package = new ExcelPackage(templateFile))
        {
            // Try to get the worksheet named "Faktury" (or "Sheet1" if you prefer)
            var worksheet = package.Workbook.Worksheets["Faktury"]
                            ?? package.Workbook.Worksheets[0];

            // We assume the first row (row 1) is for headers, so data starts from row 2
            int startRow = 2;

            // Insert data row by row
            for (int i = 0; i < payments.Count; i++)
            {
                var payment = payments[i];
                int row = startRow + i;

                // Example of filling some columns 
                // (adjust the column indices to match your template!)
                worksheet.Cells[row, 1].Value = i + 1; // Lp.
                worksheet.Cells[row, 2].Value = invoiceSufffix;
                worksheet.Cells[row, 6].Value = payment.Created.Value.ToString("yyyy-MM-dd");
                worksheet.Cells[row, 7].Value = payment.Created.Value.ToString("yyyy-MM-dd");
                worksheet.Cells[row, 8].Value = payment.Created.Value.AddDays(1).ToString("yyyy-MM-dd");
                var paymentCompanyName = payment.CompanyName.Length > 40
                    ? payment.CompanyName.Substring(0, 39)
                    : payment.CompanyName;
                if (payment.InvoiceRequested == false)
                {
                    paymentCompanyName = payment.Email;
                }
                worksheet.Cells[row, 9].Value = paymentCompanyName; // Nabywca (nazwa skrócona)
                worksheet.Cells[row, 10].Value = payment.InvoiceRequested.Value == false ? payment.Email : payment.CompanyName; // Nazwa pełna
                worksheet.Cells[row, 11].Value = payment.InvoiceRequested.Value ? payment.Street : "Paragon" ; // Adres
                worksheet.Cells[row, 12].Value = payment.InvoiceRequested.Value ? payment.PostCode : "00-000"; // Kod pocztowy
                worksheet.Cells[row, 13].Value = payment.InvoiceRequested.Value ? payment.City : "Paragon"; // Miejscowość
                worksheet.Cells[row, 14].Value = payment.Nip; // NIP
                worksheet.Cells[row, 16].Value = payment.Country; // Kraj
                worksheet.Cells[row, 17].Value = payment.InvoiceRequested.Value ? payment.Email : string.Empty; // Adres e-mail
                worksheet.Cells[row, 27].Value = "PLN"; // Waluta
                worksheet.Cells[row, 28].Value = payment.InvoiceProductName; // Nazwa towaru
                worksheet.Cells[row, 30].Value = payment.InvoiceQuantity; // Ilość
                worksheet.Cells[row, 31].Value = "szt"; // Cena jedn. netto
                worksheet.Cells[row, 33].Value = payment.Amount; // Cena jedn. netto
                worksheet.Cells[row, 34].Value = "23%"; // Stawka VAT
                worksheet.Cells[row, 35].Value = "Karta kredytowa"; // Forma płatności
                worksheet.Cells[row, 37].Value = payment.InvoiceRequested == false ? "Paragon" : string.Empty ; // Uwagi

                // If you have more columns, fill them accordingly...
            }

            // You can also format newly added rows if needed:
            int lastDataRow = startRow + payments.Count - 1;
            if (payments.Count > 0)
            {
                using (var dataRange = worksheet.Cells[startRow, 1, lastDataRow, 39])
                {
                    dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    dataRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                }
            }

            // Finally, save the result to a new file (or overwrite the templateFile if you want)
            FileInfo outputFile = new FileInfo(outputFilePath);
            package.SaveAs(outputFile);
        }
    }

    public static async Task<List<FinishedPaymentVm>> FetchPaymentsFromApiAsync(string apiUrl)
    {
        using (var httpClient = new HttpClient())
        {
            var response = await httpClient.GetStringAsync(apiUrl);
            var sessions = JsonConvert.DeserializeObject<List<PaymentSession>>(response);
            return null;
        }
    }
}