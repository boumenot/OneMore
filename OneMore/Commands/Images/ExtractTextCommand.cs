//************************************************************************************************
// Copyright © 2023 Christopher Boumenot.  All rights reserved.
//************************************************************************************************

using System.Collections.Concurrent;

namespace River.OneMoreAddIn.Commands
{
    using River.OneMoreAddIn.Models;
    using River.OneMoreAddIn.UI;

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
        }


        private async Task OcrImages(XElement[] elements)
        {
            var dict = new ConcurrentDictionary<int, string>();

            using var progress = new UI.ProgressDialog(2);
            progress.Show();
            progress.SetMessage("Extracting text from images...");

            Parallel.ForEach(
                elements.Select((x,i) => Tuple.Create(i, x)),
                new ParallelOptions { MaxDegreeOfParallelism = 10 },
                x =>
                {
                    var text = OcrImage(x.Item2).GetAwaiter().GetResult();
                    dict.AddOrUpdate(x.Item1, text, (_, s) => s);
                });

            var text = string.Join("\n\n",
                dict
                    .OrderBy(x => x.Key)
                    .Select(x => x.Value));

            progress.Increment();
            progress.SetMessage("Saving extracted text to the clipboard.");
            this.SetClipboardText(text);
            progress.Increment();
        }


        private void SetClipboardText(string text)
        {
            var thread = new Thread(() => System.Windows.Clipboard.SetText(text));
            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = false;
            thread.Start();
            thread.Join();
        }


        private async Task<string> OcrImage(XElement element)
        {
            using var image = ReadImage(element);

            string subscriptionKey = Environment.GetEnvironmentVariable("AZURE_COGNITIVE_SUBSCRIPTION_KEY");
            string endpoint = Environment.GetEnvironmentVariable("AZURE_COGNITIVE_SUBSCRIPTION_ENDPOINT");

            var client = new ComputerVisionClient(
                new ApiKeyServiceClientCredentials(subscriptionKey),
                HttpClientFactory.CreateNew(),
                disposeHttpClient: false)
            {
                Endpoint = endpoint
            };

            try
            {
                var data = Convert.FromBase64String(element.Element(ns + "Data")!.Value);
                using var stream = new MemoryStream(data);

                var result = await client.ReadInStreamAsync(stream).ConfigureAwait(false);

                var operationId = Guid.Parse(result.OperationLocation.Substring(result.OperationLocation.Length - 36));

                ReadOperationResult ror;
                do
                {
                    ror = await client.GetReadResultAsync(operationId).ConfigureAwait(false);
                    Thread.Sleep(TimeSpan.FromSeconds(1.5));
                } while (ror.Status == OperationStatusCodes.Running || ror.Status == OperationStatusCodes.NotStarted);

                var text = string.Join("\n", ror.AnalyzeResult.ReadResults
                    .SelectMany(x => x.Lines)
                    .Select(x => x.Text)
                    .ToArray());

                return text;
            }
            catch (Exception e)
            {
                this.logger.WriteVerbose(e.ToString());
                return string.Empty;
                // throw
            }
        }


        private Stream ReadImage(XElement image)
        {
            var data = Convert.FromBase64String(image.Element(ns + "Data")!.Value);
            return new MemoryStream(data, 0, data.Length);
        }
    }
}
