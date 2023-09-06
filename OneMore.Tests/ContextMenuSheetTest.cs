using Microsoft.VisualStudio.TestTools.UnitTesting;

using System.Linq;
using River.OneMoreAddIn.Settings;

namespace OneMore.Tests
{
    [TestClass]
    public class ContextMenuSheetTests
    {
        [TestMethod]
        public void HasExtractText()
        {
            var provider = new SettingsProvider();
            var testSubject = new ContextMenuSheet(provider);
            var menus = testSubject.CollectCommandMenus()
                .ToArray();

            Assert.IsTrue(menus.Any(x => x.Name == "Images Menu"));
            Assert.IsTrue(menus.Single(x => x.Name == "Images Menu").Commands.Any(x => x.Name == "Extract Text"));
        }
    }
}
