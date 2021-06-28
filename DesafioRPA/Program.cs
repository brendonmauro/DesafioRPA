using DesafioRPA.classes;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DesafioRPA
{
    class Program
    {
        private static string PathExcelListaCeps { get { return ToApplicationPath("files/Lista_de_CEPs - DESAFIO RPA.xlsx"); }  }
        private static string PathExcelResultado { get { return ToApplicationPath("files/resultado.xlsx"); } }
        private static string UrlPagCorreios { get { return "/app/cep/index.php"; } }
        private static string UrlGetDadosCorreios { get { return "https://buscacepinter.correios.com.br/app/endereco/carrega-cep-endereco.php"; } }

        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //Excel com a lista de faixas de ceps
            var excelListaCeps = new ExcelPackage(new System.IO.FileInfo(PathExcelListaCeps));
            var workSheetListaCep =  excelListaCeps.Workbook.Worksheets[0];

            // Excel que será responsável pelos resultados extraídos no sites do correios
            var excelResultado = new ExcelPackage(new System.IO.FileInfo(PathExcelResultado));
            //excelResultado.Workbook.Worksheets.Delete("Planilha1");
            //var workSheet = excelResultado.Workbook.Worksheets.Add("Planilha1");

            var workSheet = excelResultado.Workbook.Worksheets[0];

            //Definindo o cabeçalho do excel de resultado
            var indiceResultado = 1;
            workSheet.Cells[indiceResultado, 1].Value = "Logradouro";
            workSheet.Column(1).Width = 60;
            workSheet.Cells[indiceResultado, 2].Value = "Bairro";
            workSheet.Column(2).Width = 40;
            workSheet.Cells[indiceResultado, 3].Value = "Localidade/UF";
            workSheet.Column(3).Width = 40;
            workSheet.Cells[indiceResultado, 4].Value = "CEP";
            workSheet.Column(4).Width = 20;
            workSheet.Cells[indiceResultado, 5].Value = "DataProcessamento";
            workSheet.Column(5).Width = 25;

            indiceResultado++;

            var indice = 2;
            // Loop para a leitura das faixas de cep
            try
            {
                while (!string.IsNullOrEmpty(workSheetListaCep.Cells[indice, 2].Value?.ToString()))
                {
                    var numberCepInicial = Convert.ToInt64(workSheetListaCep.Cells[indice, 2].Value.ToString());
                    var numberCepFinal = Convert.ToInt64(workSheetListaCep.Cells[indice, 3].Value.ToString());

                    var currentCep = numberCepInicial;
                    HttpClient client = new HttpClient();

                    bool falhou = false;
                    int tentativasFalhas = 3;
                    //Loop para percorrer os ceps da faixa, com exceção dos últimos 3 digitos
                    while (currentCep <= numberCepFinal)
                    {
                        var formContent = new FormUrlEncodedContent(new[]
                        {
                        new KeyValuePair<string, string>("pagina", UrlPagCorreios),
                        new KeyValuePair<string, string>("endereco", (currentCep).ToString("00000000").Substring(0,5)),
                        new KeyValuePair<string, string>("tipoCEP", "ALL")
                    });

                        try
                        {
                            HttpResponseMessage response = await client.PostAsync(UrlGetDadosCorreios, formContent);
                            var stringResponse = await response.Content.ReadAsStringAsync();
                            var responseResult = JsonConvert.DeserializeObject<ResponseResult<DadosCep>>(stringResponse);

                            // caso a estrutura dos 5 digitos possua CEP ou uma lista ceps, entao é cadastrado os respectivos ceps no excel resultado
                            foreach (var dadoCep in responseResult.dados)
                            {
                                workSheet.Cells[indiceResultado, 1].Value = dadoCep.logradouroDNEC;
                                workSheet.Cells[indiceResultado, 2].Value = dadoCep.bairro;
                                workSheet.Cells[indiceResultado, 3].Value = $"{dadoCep.localidade}/{dadoCep.uf}";
                                workSheet.Cells[indiceResultado, 4].Value = dadoCep.cep;
                                workSheet.Cells[indiceResultado, 5].Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                                indiceResultado++;
                            }

                            falhou = false;
                            tentativasFalhas = 3;
                        }
                        catch (Exception e)
                        {
                            if (tentativasFalhas > 0)
                            {
                                falhou = true;
                                tentativasFalhas--;
                            }
                            else
                            {
                                workSheet.Cells[indiceResultado, 1].Value = "Conexao falhou";
                                workSheet.Cells[indiceResultado, 2].Value = $"Detalhes: {e.Message}";
                                workSheet.Cells[indiceResultado, 4].Value = (currentCep).ToString("00000000");
                                workSheet.Cells[indiceResultado, 5].Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

                                indiceResultado++;
                                falhou = false;
                                tentativasFalhas = 3;
                            }
                        }

                        if (!falhou)
                            currentCep += 1000;
                    }

                    // Salvar de 50 em 50 faixas de ceps
                    if (indice % 50 == 0)
                        excelResultado.Save();
                    indice++;
                }
            }catch { }

            excelListaCeps.Dispose();
            excelResultado.Save();
            excelResultado.Dispose();
        }

        public static string ToApplicationPath(string fileName)
        {
            var exePath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
            Regex appPathMatcher = new Regex(@"(?<!fil)[A-Za-z]:\\+[\S\s]*?(?=\\+bin)");
            var appRoot = appPathMatcher.Match(exePath).Value;
            return Path.Combine(appRoot, fileName);
        }
    }
}
