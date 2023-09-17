//************************************************************************************************
// Copyright © 2023 Christopher Boumenot.  All rights reserved.
//************************************************************************************************

namespace River.OneMoreAddIn.Commands
{
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Xml.Linq;


    /// <summary>
    /// Unwraps a paragraph by removing hard coded line breaks in the text.  OneNote automatically
    /// flows text.  There is no need to have explicit new lines.
    /// </summary>
    internal class UnwrapCommand : Command
    {
        public UnwrapCommand()
        {
        }


        public override async Task Execute(params object[] args)
        {
            using var one = new OneNote(out var page, out _);

            var selections = page.GetSelectedElements().ToArray();
            if (selections.Length < 2)
            {
                return;
            }
            
            var sb = new StringBuilder();
            for (int i=0; i < selections.Length; i++)
            {
                var selection = selections[i];
                var wrapper = selection.GetCData().GetWrapper();

                var lines = wrapper
                    .DescendantNodes()
                    .OfType<XText>()
                    .Select(x => x.Value.Trim());

                var text = string.Join(" ", lines);
                sb.Append(text);

                if (i > 0)
                {
                    selection.Parent.Remove();
                }

                if (i != selections.Length - 1)
                {
                    sb.Append(" ");
                }
            }

            selections[0].FirstNode.ReplaceWith(new XCData(sb.ToString()));
            await one.Update(page);

            logger.WriteLine($"unwrapped {selections.Length} lines");
        }
    }
}
