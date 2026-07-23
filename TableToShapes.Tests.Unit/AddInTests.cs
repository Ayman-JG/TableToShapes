using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using FluentAssertions;
using NUnit.Framework;
using TableToShapes.AddIn;

namespace TableToShapes.Tests.Unit
{
    /// <summary>
    /// Tests the add-in's static surface that does not need a running PowerPoint: the ribbon
    /// definition and the COM registration attributes (which install-addin.ps1 depends on).
    /// COM activation, selection handling and the message boxes remain integration/manual - the
    /// conversion pipeline the button calls is covered by the E2E tests via TableConverter.
    /// </summary>
    [TestFixture]
    public class AddInTests
    {
        // These must match install-addin.ps1 exactly, or registration/discovery breaks.
        private const string ExpectedClsid = "6F9B0C64-3C1A-4E6E-9B7D-2D3E8A11F0AB";
        private const string ExpectedProgId = "TableToShapes.AddIn.Connect";
        private const string RibbonOnAction = "OnConvertClicked";

        [Test]
        public void GetCustomUI_ReturnsWellFormedRibbonWithTheConvertButton()
        {
            var xml = new Connect().GetCustomUI(null);

            xml.Should().NotBeNullOrWhiteSpace();
            Action parse = () => XDocument.Parse(xml);
            parse.Should().NotThrow("the ribbon XML must be well-formed or the add-in fails to load");

            xml.Should().Contain("idMso='TabHome'");
            xml.Should().Contain("id='ConvertTableButton'");
            xml.Should().Contain("onAction='" + RibbonOnAction + "'");
        }

        [Test]
        public void RibbonOnActionCallback_ExistsAsAPublicMethod()
        {
            // The ribbon resolves the callback by name at runtime, so a rename silently breaks it.
            typeof(Connect).GetMethod(RibbonOnAction).Should().NotBeNull();
        }

        [Test]
        public void ComRegistrationAttributes_MatchTheInstallScript()
        {
            var type = typeof(Connect);

            type.GUID.Should().Be(new Guid(ExpectedClsid));

            var progId = (ProgIdAttribute)Attribute.GetCustomAttribute(type, typeof(ProgIdAttribute));
            progId.Should().NotBeNull();
            progId.Value.Should().Be(ExpectedProgId);

            var comVisible = (ComVisibleAttribute)Attribute.GetCustomAttribute(type, typeof(ComVisibleAttribute));
            comVisible.Should().NotBeNull();
            comVisible.Value.Should().BeTrue();
        }
    }
}
