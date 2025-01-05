namespace ITBees.SaldeoIntegration
{
    using System;
    using System.IO;
    using System.IO.Compression;
    using System.Net.Http;
    using System.Security.Cryptography;
    using System.Text;
    using System.Threading.Tasks;
    using System.Web;

    public class SaldeoInvoice
    {
        private const string username = "TWOJ_LOGIN";
        private const string apiToken = "TWÓJ_API_TOKEN";
        private const string companyProgramId = "ABC123";
        private static readonly string saldeoBaseUrl = "https://saldeo.brainshare.pl";
        // lub np. "https://saldeo-test.brainshare.pl" – zależnie od środowiska

        // -------------------------------------------------------------------
        // 1) ComputeSignature - calculates req_sig for Saldeo
        // -------------------------------------------------------------------
        /// <summary>
        /// Computes the Saldeo request signature (req_sig).
        /// According to Saldeo docs, you need to:
        /// 1. Sort parameters alphabetically by their key (excluding empty ones).
        /// 2. Build a string in form "key1=value1key2=value2..." (no &).
        /// 3. Apply URL encoding for the entire string.
        /// 4. Append the api_token.
        /// 5. Calculate the MD5 of the result and convert to HEX.
        /// </summary>
        private static string ComputeSignature(params (string key, string value)[] parameters)
        {
            // 1) Sort parameters by key
            Array.Sort(parameters, (a, b) => string.Compare(a.key, b.key, StringComparison.Ordinal));

            // 2) Build raw query (concatenated "key=value" pairs) 
            var concatenated = new StringBuilder();
            foreach (var param in parameters)
            {
                // Skip empty values
                if (string.IsNullOrEmpty(param.value)) continue;

                concatenated.Append(param.key).Append("=").Append(param.value);
            }

            // 3) URL encode result - notice the custom rules in Saldeo doc 
            // (space -> '+', uppercase HEX after %)
            string urlEncoded = CustomUrlEncode(concatenated.ToString());

            // 4) Append apiToken
            string toSign = urlEncoded + apiToken;

            // 5) MD5 -> HEX
            using (var md5 = MD5.Create())
            {
                var hash = md5.ComputeHash(Encoding.UTF8.GetBytes(toSign));
                return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Minimal custom URL-encoding that tries to follow Saldeo's rules
        /// (the doc shows some differences from standard .NET UrlEncode).
        /// For simplicity, here we use standard .NET plus uppercase fix. 
        /// In real code, you'll want to adapt to the exact rules from the doc.
        /// </summary>
        private static string CustomUrlEncode(string raw)
        {
            // .NET's HttpUtility.UrlEncode might produce lowercase hex and %20 for space, etc.
            // We'll do a quick workaround for uppercase and space->'+':
            string standard = HttpUtility.UrlEncode(raw, Encoding.UTF8);

            if (standard == null) return string.Empty;

            // Replace %20 with +, and also uppercase any %xx
            // This is a partial approach. See doc for full compliance if needed.
            standard = standard.Replace("%20", "+");
            // Make %xx uppercase
            // (In real code, you'd parse carefully. Here we do a naive approach.)
            var sb = new StringBuilder();
            for (int i = 0; i < standard.Length; i++)
            {
                if (standard[i] == '%' && i + 2 < standard.Length)
                {
                    sb.Append('%');
                    sb.Append(char.ToUpper(standard[i + 1]));
                    sb.Append(char.ToUpper(standard[i + 2]));
                    i += 2;
                }
                else
                {
                    sb.Append(standard[i]);
                }
            }
            return sb.ToString();
        }

        // -------------------------------------------------------------------
        // 2) BuildAddInvoiceXml - constructs example XML body for "document.add"
        // -------------------------------------------------------------------
        /// <summary>
        /// Builds an example XML (uncompressed) with one "document" that will be recognized as an invoice.
        /// In real usage, you must follow Saldeo's XSD for "document/add".
        /// This is a heavily simplified snippet.
        /// </summary>
        private static string BuildAddInvoiceXml()
        {
            // Example "document.add" request (simplified).
            // Real payload must match Saldeo's XSD carefully: 
            // https://saldeo.brainshare.pl/static/doc/api/document/document_add_request.xml
            return
    @"<?xml version=""1.0"" encoding=""UTF-8""?>
<ROOT>
  <DOCUMENTS>
    <DOCUMENT>
      <YEAR>2024</YEAR>
      <MONTH>08</MONTH>
      <FILENAME>faktura_sprzedazowa.pdf</FILENAME>
      <!-- Saldeo needs an attachment reference matching param attmnt_id below -->
      <ATTMNT_ID>file1</ATTMNT_ID>
      <!-- 
        Additional fields:
        <DOC_TYPE>FV</DOC_TYPE> or some code to mark it's a sales invoice, 
        <GUID>... if needed ...</GUID>, etc.
      -->
    </DOCUMENT>
  </DOCUMENTS>
</ROOT>";
        }

        // -------------------------------------------------------------------
        // 3) GzipAndBase64Encode - gzips the XML and base64-encodes it
        // -------------------------------------------------------------------
        /// <summary>
        /// Takes raw XML, gzips it, and returns Base64 string for Saldeo's "command" param.
        /// </summary>
        private static string GzipAndBase64Encode(string xml)
        {
            byte[] rawBytes = Encoding.UTF8.GetBytes(xml);

            using (var output = new MemoryStream())
            {
                using (var gzip = new GZipStream(output, CompressionMode.Compress))
                {
                    gzip.Write(rawBytes, 0, rawBytes.Length);
                }
                return Convert.ToBase64String(output.ToArray());
            }
        }

        // -------------------------------------------------------------------
        // 4) AddInvoiceAsync - sends POST to Saldeo for "document.add"
        // -------------------------------------------------------------------
        /// <summary>
        /// Sends POST to Saldeo to add a new sales invoice (treated as a document).
        /// Also sends the PDF file as attachment (example).
        /// In reality, you must adapt to your actual model and fill the correct fields.
        /// </summary>
        public static async Task AddInvoiceAsync(string pdfFilePath)
        {
            // 1) Build uncompressed XML
            string xmlToSend = BuildAddInvoiceXml();

            // 2) GZIP + Base64
            string commandGzipped = GzipAndBase64Encode(xmlToSend);

            // 3) Prepare Saldeo required parameters
            string reqId = DateTime.Now.ToString("yyyyMMddHHmmssfff"); // your unique request ID
                                                                       // You could store them in a list to pass to signature
            var signatureParams = new (string key, string value)[]
            {
            ("req_id", reqId),
            ("username", username)
            };
            string reqSig = ComputeSignature(signatureParams);

            // 4) Prepare form-data with "command" and "attmnt_file1"
            //    Saldeo docs: param name for attachments is "attmnt_<ID>" -> attmnt_file1
            var fileBytes = File.ReadAllBytes(pdfFilePath);
            var fileGzippedBase64 = GzipAndBase64Encode(Encoding.UTF8.GetString(fileBytes));
            // Real usage: send the raw PDF in GZip+Base64. 
            // If you want to keep filename, it can be stored in FILENAME node or "title" param.

            using (var httpClient = new HttpClient())
            {
                // Endpoint e.g.: https://saldeo.brainshare.pl/api/xml/1.0/document/add
                var url = $"{saldeoBaseUrl}/api/xml/1.0/document/add";

                using (var multipartContent = new MultipartFormDataContent())
                {
                    multipartContent.Add(new StringContent(username), "username");
                    multipartContent.Add(new StringContent(reqId), "req_id");
                    multipartContent.Add(new StringContent(reqSig), "req_sig");
                    multipartContent.Add(new StringContent(companyProgramId), "company_program_id");

                    // XML in param "command"
                    multipartContent.Add(new StringContent(commandGzipped), "command");

                    // Attachment -> must match <ATTMNT_ID>file1</ATTMNT_ID> from XML
                    // param name: attmnt_file1
                    multipartContent.Add(new StringContent(fileGzippedBase64), "attmnt_file1");

                    // Send POST
                    var response = await httpClient.PostAsync(url, multipartContent);
                    var responseData = await response.Content.ReadAsStringAsync();

                    // For debugging
                    Console.WriteLine("AddInvoiceAsync response:");
                    Console.WriteLine(responseData);
                }
            }
        }

        // -------------------------------------------------------------------
        // 5) GetInvoicePdfAsync - example of retrieving PDF link
        // -------------------------------------------------------------------
        /// <summary>
        /// Example of how to read PDF link for the newly added invoice by calling e.g. `invoice/list` or `document/list`.
        /// Then we download the file. In a real scenario, adapt to your logic (matching GUID/ID, etc.).
        /// </summary>
        public static async Task GetInvoicePdfAsync(string targetInvoiceNumber, string savePath)
        {
            // 1) Build the GET call for invoice/list (or document.list).
            // E.g.: GET /api/xml/1.8/invoice/list?company_program_id=XYZ&policy=SALDEO&username=...
            // with the same approach for req_sig.

            string reqId = DateTime.Now.ToString("yyyyMMddHHmmssfff");

            // Sort by key: company_program_id, policy, req_id, username
            var signatureParams = new (string key, string value)[]
            {
            ("company_program_id", companyProgramId),
            ("policy", "SALDEO"),
            ("req_id", reqId),
            ("username", username)
            };
            string reqSig = ComputeSignature(signatureParams);

            // Example: 
            // GET /api/xml/1.8/invoice/list?company_program_id=companyProgramId&policy=SALDEO&username=...&req_id=...&req_sig=...
            var url = $"{saldeoBaseUrl}/api/xml/1.8/invoice/list" +
                      $"?company_program_id={companyProgramId}" +
                      $"&policy=SALDEO" +
                      $"&username={username}" +
                      $"&req_id={reqId}" +
                      $"&req_sig={reqSig}";

            using (var httpClient = new HttpClient())
            {
                var response = await httpClient.GetAsync(url);
                var xmlResponse = await response.Content.ReadAsStringAsync();

                // TODO: parse XML to find the invoice matching e.g. <NUMBER>someNumber</NUMBER>
                // and read the <SOURCE> or <PREVIEW_URL> link to PDF.

                // Suppose we found something like:
                string pdfLink = "https://saldeo.brainshare.pl/.../some-random-file.pdf?todownload=Invoice123.pdf";

                // 2) Download PDF from that link
                var pdfBytes = await httpClient.GetByteArrayAsync(pdfLink);

                // 3) Save locally
                File.WriteAllBytes(savePath, pdfBytes);

                Console.WriteLine("Saved invoice PDF to: " + savePath);
            }
        }

        // -------------------------------------------------------------------
        // Example usage
        // -------------------------------------------------------------------
        public static async Task Main()
        {
            // 1) Add invoice (document) with a PDF file
            string localPdfPath = @"C:\faktura.pdf";
            await AddInvoiceAsync(localPdfPath);

            // 2) Optionally, fetch the PDF link from invoice/list (if you want to confirm or re-download)
            await GetInvoicePdfAsync("FV-123/2025", @"C:\temp\downloaded-invoice.pdf");

            Console.WriteLine("Done.");
            Console.ReadKey();
        }
    }
}
