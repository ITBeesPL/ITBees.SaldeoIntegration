using ITBees.FAS.Payments.Interfaces.Models;
using Newtonsoft.Json;
using OfficeOpenXml;

namespace ITBees.SaldeoIntegration;

public class ExcelGenerator
{
    public static void GenerateExcelReport(List<FinishedPaymentVm> payments, string filePath)
    {
        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Faktury");

            // Headers (Polish column names)
            worksheet.Cells[1, 1].Value = "Lp.";
            worksheet.Cells[1, 2].Value = "Sufiks";
            worksheet.Cells[1, 3].Value = "Własny opis faktury";
            worksheet.Cells[1, 4].Value = "Metoda kasowa";
            worksheet.Cells[1, 5].Value = "VAT marża";
            worksheet.Cells[1, 6].Value = "Data wystawienia";
            worksheet.Cells[1, 7].Value = "Data dostawy";
            worksheet.Cells[1, 8].Value = "Termin płatności";
            worksheet.Cells[1, 9].Value = "Nabywca (nazwa skrócona kontrahenta)";
            worksheet.Cells[1, 10].Value = "Nazwa pełna kontrahenta";
            worksheet.Cells[1, 11].Value = "Adres";
            worksheet.Cells[1, 12].Value = "Kod pocztowy";
            worksheet.Cells[1, 13].Value = "Miejscowość";
            worksheet.Cells[1, 14].Value = "NIP";
            worksheet.Cells[1, 15].Value = "REGON";
            worksheet.Cells[1, 16].Value = "Kraj - skrót kraju w formacie ISO";
            worksheet.Cells[1, 17].Value = "Adres e-mail";
            worksheet.Cells[1, 18].Value = "Odbiorca (nazwa skrócona kontrahenta)";
            worksheet.Cells[1, 19].Value = "Nazwa pełna (odbiorcy)";
            worksheet.Cells[1, 20].Value = "Adres (odbiorcy)";
            worksheet.Cells[1, 21].Value = "Kod pocztowy (odbiorcy)";
            worksheet.Cells[1, 22].Value = "Miejscowość (odbiorcy)";
            worksheet.Cells[1, 23].Value = "NIP (odbiorcy)";
            worksheet.Cells[1, 24].Value = "REGON (odbiorcy)";
            worksheet.Cells[1, 25].Value = "Kraj - skrót kraju w formacie ISO (odbiorcy)";
            worksheet.Cells[1, 26].Value = "Adres e-mail (odbiorcy)";
            worksheet.Cells[1, 27].Value = "Waluta";
            worksheet.Cells[1, 28].Value = "Nazwa towaru";
            worksheet.Cells[1, 29].Value = "PKWiU";
            worksheet.Cells[1, 30].Value = "Ilość";
            worksheet.Cells[1, 31].Value = "Jednostka";
            worksheet.Cells[1, 32].Value = "Cena jedn. netto";
            worksheet.Cells[1, 33].Value = "Cena jedn. brutto";
            worksheet.Cells[1, 34].Value = "Stawka VAT";
            worksheet.Cells[1, 35].Value = "Forma płatności";
            worksheet.Cells[1, 36].Value = "Konto bankowe";
            worksheet.Cells[1, 37].Value = "Uwagi";
            worksheet.Cells[1, 38].Value = "Wystawca faktury";
            worksheet.Cells[1, 39].Value = "Grupa Towarowa";

            // Data rows
            for (int i = 0; i < payments.Count; i++)
            {
                var payment = payments[i];
                int row = i + 2;
                worksheet.Cells[row, 1].Value = i + 1; // Lp.
                worksheet.Cells[row, 9].Value = payment.CompanyName; // Nabywca (nazwa skrócona kontrahenta)
                worksheet.Cells[row, 10].Value = payment.CompanyName; // Nazwa pełna kontrahenta
                worksheet.Cells[row, 11].Value = payment.Street; // Adres
                worksheet.Cells[row, 12].Value = payment.PostCode; // Kod pocztowy
                worksheet.Cells[row, 13].Value = payment.City; // Miejscowość
                worksheet.Cells[row, 14].Value = payment.Nip; // NIP
                worksheet.Cells[row, 16].Value = payment.Country; // Kraj - skrót kraju w formacie ISO
                worksheet.Cells[row, 17].Value = payment.Email; // Adres e-mail
                worksheet.Cells[row, 27].Value = "PLN"; // Waluta
                worksheet.Cells[row, 28].Value = payment.InvoiceProductName; // Nazwa towaru
                worksheet.Cells[row, 30].Value = payment.InvoiceQuantity; // Ilość
                worksheet.Cells[row, 32].Value = payment.Amount; // Cena jedn. netto
                worksheet.Cells[row, 34].Value = "23%"; // Stawka VAT
                worksheet.Cells[row, 35].Value = "Przelew"; // Forma płatności
            }

            // Auto fit columns
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

            // Save to file
            package.SaveAs(new FileInfo(filePath));
        }
    }

    public static async Task<List<FinishedPaymentVm>> FetchPaymentsFromApiAsync(string apiUrl)
    {
        using (var httpClient = new HttpClient())
        {
            var response = await httpClient.GetStringAsync(apiUrl);
            var sessions = JsonConvert.DeserializeObject<List<PaymentSession>>(response);
            return FinishedPaymentVm.MapFromPaymentSessions(sessions);
        }
    }
}