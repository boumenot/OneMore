namespace River.OneMoreAddIn.Commands
{
    using River.OneMoreAddIn.Settings;
    using River.OneMoreAddIn.Models;

    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;
    using System.Xml.Linq;

    using Resx = Properties.Resources;
    using Microsoft.Azure.CognitiveServices.Vision.ComputerVision;
    using Microsoft.Azure.CognitiveServices.Vision.ComputerVision.Models;
    using System.Threading;

    internal class ExtractTextCommand : Command
    {
        private Page page;
        private XNamespace ns;

        public ExtractTextCommand()
        {
        }

        public override async Task Execute(params object[] args)
        {
            if (!HttpClientFactory.IsNetworkAvailable())
            {
                UIHelper.ShowInfo(Resx.NetwordConnectionUnavailable);
                return;
            }

            using var one = new OneNote(out this.page, out this.ns, OneNote.PageDetail.All);
            var elements = page.Root.Descendants(ns + "Image")?
                .Where(e => e.Attribute("selected")?.Value == "all")
                .ToArray();

            if (!elements.Any())
            {
                // starting at Outline should exclude all background images
                elements = page.Root.Elements(ns + "Outline").Descendants(ns + "Image").ToArray();
            }

            if (elements.Any())
            {
                await OcrImages(elements);
            }
            else
            {
                UIHelper.ShowMessage(Resx.ResizeImagesDialog_noImages);
            }

            await Task.FromResult(0);
        }

        private async Task OcrImages(IEnumerable<XElement> elements)
        {
            foreach (var element in elements)
            {
                await OcrImage(element);
            }
        }

        private async Task OcrImage(XElement element)
        {
            using var image = ReadImage(element);

            string subscriptionKey = Environment.GetEnvironmentVariable("AZURE_COGNITIVE_SUBSCRIPTION_KEY");
            string endpoint = "https://me.cognitiveservices.azure.com/";

            var client = new ComputerVisionClient(new ApiKeyServiceClientCredentials(subscriptionKey)) { Endpoint = endpoint };

            var data = Convert.FromBase64String(element.Element(ns + "Data").Value);
            using var stream = new MemoryStream(data, 0, data.Length);

            var result = await client.ReadInStreamAsync(stream);

            var operationId = Guid.Parse(result.OperationLocation.Substring(result.OperationLocation.Length - 36));

            ReadOperationResult ror;
            do
            {
                ror = await client.GetReadResultAsync(operationId);
                Thread.Sleep(TimeSpan.FromSeconds(3));
            } while (ror.Status == OperationStatusCodes.Running || ror.Status == OperationStatusCodes.NotStarted);

            var text = string.Join("\n", ror.AnalyzeResult.ReadResults
                .SelectMany(x => x.Lines)
                .Select(x => x.Text)
                .ToArray());
        }

        private Stream ReadImage(XElement image)
        {
            var data = Convert.FromBase64String(image.Element(ns + "Data").Value);
            return new MemoryStream(data, 0, data.Length);
        }
    }
}
