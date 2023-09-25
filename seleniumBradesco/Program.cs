using System;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Support.UI;
using System.Threading;
using System.Drawing;
using OpenQA.Selenium.Interactions;
using Newtonsoft.Json;
using System;
using System.Net.Http;
using System.Threading;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Newtonsoft.Json.Linq;
using System.Net;
using System.Text;
using System.Net.Http.Headers;
using OpenQA.Selenium.Edge;


namespace IEDriverSample
{
    class Program
    {
        static async Task Main(string[] args)
        {
            string authToken = "F51778AD-C1D3-4CDD-AE3E-825AA3281991";
            string jsonUrl = "https://ppintegracaoapi.azurewebsites.net/RoboBradesco/AtasDisponiveis";

            var ieOptions = new InternetExplorerOptions();
            ieOptions.AttachToEdgeChrome = true;
            ieOptions.EdgeExecutablePath = "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe";

            var driver = new InternetExplorerDriver(ieOptions);

            // Navegar para o site
            //driver.Navigate().GoToUrl("https://jte.csjt.jus.br/");

            //IWebDriver driver = new InternetExplorerDriver();

            HashSet<int> idsProcessados = new HashSet<int>(); // Conjunto para armazenar os IDs processados

            try
            {
                string jsonContent = await GetJsonContent(jsonUrl, authToken);

                List<Tramitacao> tramitacoes = JsonConvert.DeserializeObject<List<Tramitacao>>(jsonContent);

                driver.Url = "https://juridico8.bradesco.com.br/gcpj/menuFrames.htm";

                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                wait.Until(driver => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));

                IWebElement frameCorpoINET = driver.FindElement(By.Name("frameCorpoINET"));
                driver.SwitchTo().Frame(frameCorpoINET);

                By btnLoginSelector = By.CssSelector("#btnLogin");
                IWebElement btnLogin = WaitUntilVisible(driver, btnLoginSelector);
                btnLogin.Click();

                bool attProcessosVisible = false;
                int maxAttempts = 25; // Número máximo de tentativas
                int currentAttempt = 0;
                while (!attProcessosVisible && currentAttempt < maxAttempts)
                {
                    try
                    {
                        IWebElement attProcessos = driver.FindElement(By.LinkText("Atualização do Andamento dos Processos"));
                        if (attProcessos.Displayed)
                        {
                            attProcessosVisible = true;
                            attProcessos.Click();
                        }
                    }
                    catch (NoSuchElementException)
                    {
                        // Elemento não encontrado, esperar e tentar novamente
                        Thread.Sleep(1000); // Aguardar 1 segundo
                        currentAttempt++;
                    }
                }

                foreach (var tramitacao in tramitacoes)
                {
                    int idTramitacao = tramitacao.id_tramitacao;

                    // Verificar se o ID já foi processado
                    if (idsProcessados.Contains(idTramitacao))
                    {
                        Console.WriteLine($"ID {idTramitacao} já foi processado. Pulando para o próximo.");
                        continue;
                    }

                    string numeroIntegracao = tramitacao.numero_integracao;
                    string textoOpcao = tramitacao.evento_bradesco;
                    int index = textoOpcao.IndexOf(">");
                    if (index >= 0 && index + 1 < textoOpcao.Length)
                    {
                        textoOpcao = textoOpcao.Substring(index + 1).Trim();
                    }
                    DownloadPdfFromUrl(tramitacao.texto, @"C:\Bradesco\Ata\AtaAudiencia_" + tramitacao.numero_integracao + ".pdf");

                    //Numero processo bradesco e clicar para pesquisar
                    IWebElement campoNumero = driver.FindElement(By.Name("cdNumeroProcessoBradesco"));
                    campoNumero.Clear();
                    campoNumero.SendKeys(numeroIntegracao);
                    IWebElement btnPesquisar = driver.FindElement(By.Name("btoPesquisar"));
                    btnPesquisar.Click();

                    if (IsAlertPresent(driver))
                    {
                        IAlert alertErro = driver.SwitchTo().Alert();
                        string mensagemErro = alertErro.Text;
                        Console.WriteLine("Texto do alerta de erro de pesquisa: " + mensagemErro);
                        alertErro.Accept();

                        await EnviarDadosParaLink(idTramitacao, mensagemErro, authToken);

                        Thread.Sleep(3000);

                        continue;
                    }
                    

                    Thread.Sleep(3000);

                    //Select de tipo de audiencia
                    IWebElement selectElement = driver.FindElement(By.Name("cdReferenciaAndamentoProcesso"));
                    SelectElement select = new SelectElement(selectElement);
                    select.SelectByText(textoOpcao);

                    //Texto de ata
                    IWebElement textareaElement = driver.FindElement(By.Id("dsAndamentoProcessoEscritorio"));
                    textareaElement.Clear();
                    
                    int chunkSize = 1000; // Number of characters per chunk
                    for (int i = 0; i < tramitacao.conteudo_ata.Length; i += chunkSize)
                    {
                        string chunk = tramitacao.conteudo_ata.Substring(i, Math.Min(chunkSize, tramitacao.conteudo_ata.Length - i));
                        textareaElement.SendKeys(chunk);
                    }

                    //Entrar em anexos
                    IWebElement anexos = driver.FindElement(By.Name("btoAnexos"));
                    anexos.Click();

                    Thread.Sleep(5000);

                    IWebElement iframe = driver.FindElement(By.Id("here"));
                    driver.SwitchTo().Frame(iframe);
                    //Nome do anexo
                    IWebElement campoNomeAnexo = driver.FindElement(By.Name("nmAnexoProcesso"));
                    string nomeAnexo = "ATA - " + numeroIntegracao;
                    campoNomeAnexo.SendKeys(nomeAnexo);

                    //Campo select do anexo
                    IWebElement selectAnexos = driver.FindElement(By.Name("cmdAnexos"));
                    SelectElement selectAnexo = new SelectElement(selectAnexos);
                    selectAnexo.SelectByText("TR - ATA DE AUDIENCIA");


                    string caminhoArquivo = @"C:\Bradesco\Ata\AtaAudiencia_" + numeroIntegracao + ".pdf";
                    IWebElement inputFile = driver.FindElement(By.Name("formFile"));
                    inputFile.SendKeys(caminhoArquivo);

                    IWebElement incluirAta = driver.FindElement(By.Name("btoIncluir"));
                    incluirAta.Click();

                    IWebElement voltar = driver.FindElement(By.Name("btoVoltar"));
                    voltar.Click();

                    Thread.Sleep(3000);

                    driver.SwitchTo().DefaultContent();
                    driver.SwitchTo().Frame(frameCorpoINET);

                    //< input name = "btoSalvar" class="bto1" style="cursor: hand;" onmouseover="this.style.cursor='hand'" onclick="enviar( 'execucaoSalvar', true );" type="button" value="salvar">

                    Console.WriteLine("Numero processo bradesco:" + numeroIntegracao + ", Id tramitacao:" + idTramitacao);

                    //< input name = "btoSalvar" class="bto1" style="cursor: hand;" onmouseover="this.style.cursor='hand'" onclick="enviar( 'execucaoSalvar', true );" type="button" value="salvar">

                    IWebElement salvar = driver.FindElement(By.Name("btoSalvar"));
                    salvar.Click();

                    Thread.Sleep(2000);

                    // Alternar para o alerta/popup
                    IAlert alert = driver.SwitchTo().Alert();

                    // Obter o texto do alerta/popup
                    string alertText = alert.Text;
                    alertText = alert.Text;
                    Console.WriteLine("Texto do popup/alerta: " + alertText);
                    // Aceitar o alerta/popup (ou realizar outras ações necessárias)
                    alert.Accept();



                    //< input name = "mensagem" type = "hidden" value = "Operação realizada com sucesso" >

                    // Fazer o post
                    await EnviarDadosParaLink(idTramitacao, alertText, authToken);

                    Thread.Sleep(3000);

                    

                    // Adicionar o ID ao conjunto de IDs processados após o processamento
                    idsProcessados.Add(idTramitacao);
                }
                driver.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ocorreu um erro: " + ex.Message);
            }
            finally
            {
                driver.Quit();
            }
        }
        static bool IsAlertPresent(IWebDriver driver)
        {
            try
            {
                driver.SwitchTo().Alert();
                return true;
            }
            catch (NoAlertPresentException)
            {
                return false;
            }
        }


        static IWebElement WaitUntilVisible(IWebDriver driver, By selector, int timeoutInSeconds = 10)
        {
            DateTime endTime = DateTime.Now.AddSeconds(timeoutInSeconds);
            while (DateTime.Now < endTime)
            {
                try
                {
                    var element = driver.FindElement(selector);
                    if (element.Displayed)
                        return element;
                }
                catch (NoSuchElementException) { }
                catch (StaleElementReferenceException) { }

                Thread.Sleep(500);
            }

            throw new NoSuchElementException($"Elemento {selector} não encontrado ou não está visível após {timeoutInSeconds} segundos.");
        }
        public static void DownloadPdfFromUrl(string url, string fileName)
        {
            using (WebClient wc = new WebClient())
            {
                wc.DownloadFile(url, fileName);
            }
        }

        public static string ReadPdfFromUrl(string url)
        {
            WebClient wc = new WebClient();
            byte[] bytes = wc.DownloadData(url);
            PdfReader reader = new PdfReader(bytes);
            string text = string.Empty;
            for (int page = 1; page <= reader.NumberOfPages; page++)
            {
                text += PdfTextExtractor.GetTextFromPage(reader, page);
            }
            reader.Close();
            return text;
        }
        // Classes do JSON
        public class Tramitacao
        {
            public int id_tramitacao { get; set; }
            public string evento { get; set; }
            public string evento_bradesco { get; set; }
            public DateTime ag_data_hora { get; set; }
            public string numero_processo { get; set; }
            public string texto { get; set; }
            public string conteudo_ata { get; set; }
            public string numero_integracao { get; set; }
            public object id { get; set; }
        }

        //GET
        static async Task<string> GetJsonContent(string url, string authToken)
        {
            using (HttpClient client = new HttpClient())
            {
                // Set the custom authorization header with the token
                client.DefaultRequestHeaders.Add("Authorization", authToken);

                HttpResponseMessage response = await client.GetAsync(url);
                if (response.IsSuccessStatusCode)
                {
                    return await response.Content.ReadAsStringAsync();
                }
                else
                {
                    throw new Exception("Error fetching JSON. Status code: " + response.StatusCode);
                }
            }
        }

        //POST
        static async Task EnviarDadosParaLink(int idTramitacao, string mensagem, string authToken)
        {
            string apiUrl = "https://ppintegracaoapi.azurewebsites.net/RoboBradesco/AtaProcessada";

            var postData = new
            {
                id = idTramitacao,
                mensagem = mensagem
            };

            using (HttpClient client = new HttpClient())
            {
                // Set the custom authorization header with the token
                client.DefaultRequestHeaders.Add("Authorization", authToken);

                var jsonContent = new StringContent(JsonConvert.SerializeObject(postData), Encoding.UTF8, "application/json");

                // Send the POST request
                HttpResponseMessage response = await client.PostAsync(apiUrl, jsonContent);

                // Check the response
                if (response.IsSuccessStatusCode)
                {
                    string responseBody = await response.Content.ReadAsStringAsync();
                    Console.WriteLine("Resposta do POST: " + responseBody);
                }
                else
                {
                    Console.WriteLine(response.StatusCode);
                }
            }
        }


    }
}
