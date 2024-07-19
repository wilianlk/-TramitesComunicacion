using System;
using System.Threading.Tasks;
using RestSharp;
using Newtonsoft.Json;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System.Collections.Generic;
using System.Linq;

namespace TramitesComunicacion.ViewModels
{
    internal class WebServiceClient
    {
        private readonly RestClient client;
        private void ConfigureRequest(RestRequest request)
        {
            request.AddHeader("Content-Type", "application/json");
            string token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJDQzE2ODQyMzA0IiwianRpIjoiNzMwNjAiLCJyb2xlcyI6IlBST1ZFRURPUl9ERV9CSUVORVNfWV9TRVJWSUNJT1MiLCJpZEVtcHJlc2EiOjEwNzczMSwiaWRNb2R1bG8iOjQzMywicGVydGVuZWNlQSI6IkRERiIsInR5cGUiOiJleHRlcm5hbCIsIm5vbWJyZU1vZHVsbyI6IlJORSBMZXkgRGVqZW4gZGUgRnJlZ2FyIiwiaWF0IjoxNzIxMjIwODQ3LCJleHAiOjE3MzY5ODg4NDd9.zCVsHQuSzIwymJJ-6UmqujnIeSYSLHsJdZry92jFCAQ";
            request.AddHeader("Authorization", $"Bearer {token}");
            request.AddHeader("Cookie", "cookiesession1=YOUR_COOKIE_HERE"); // Consider handling cookies more securely
        }
        public WebServiceClient()
        {
            var options = new RestClientOptions("https://tramitescrcom.gov.co")
            {
                MaxTimeout = 10000 
            };
            client = new RestClient(options);
        }
        public async Task<string> ConsultarRnePorCorreoAsync(string[] emails)
        {
            var request = new RestRequest("/excluidosback/consultaMasiva/validarExcluidos", Method.Post);
            ConfigureRequest(request);
            request.AddJsonBody(new { type = "COR", keys = emails });

            return await ExecuteRequestAsync(request);
        }
        public async Task<string> ConsultarRnePorTelefonoAsync(string[] telefonos)
        {
            var request = new RestRequest("/excluidosback/consultaMasiva/validarExcluidos", Method.Post);
            ConfigureRequest(request);
            request.AddJsonBody(new { type = "TEL", keys = telefonos });

            return await ExecuteRequestAsync(request);
        }
        private async Task<string> ExecuteRequestAsync(RestRequest request)
        {
            try
            {
                RestResponse response = await client.ExecuteAsync(request);
                if (!response.IsSuccessful)
                {
                    return $"Error: {response.StatusCode} - {response.ErrorMessage}";
                }

                var resultados = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(response.Content);
                int countRespuestas = resultados.Count(item => item.ContainsKey("respuesta") && item["respuesta"] != null);

                Console.WriteLine($"Total respuestas devueltas por api: {countRespuestas} de {resultados.Count}");

                return response.Content;
            }
            catch (Exception ex)
            {
                return $"Exception occurred: {ex.Message}";
            }
        }
    }
}
