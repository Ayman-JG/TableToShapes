using System;
using System.IO;
using FluentAssertions;
using NUnit.Framework;
using TableToShapes.Core.Logging;

namespace TableToShapes.Tests.Unit
{
    [TestFixture]
    public class LogLevelParserTests
    {
        [TestCase("debug", LogLevel.Debug)]
        [TestCase("INFO", LogLevel.Info)]
        [TestCase("Warning", LogLevel.Warning)]
        [TestCase("error", LogLevel.Error)]
        [TestCase("none", LogLevel.None)]
        public void GivenValidName_WhenParsing_ThenReturnsThatLevel(string input, LogLevel expected)
            => LogLevelParser.Parse(input).Should().Be(expected);

        [TestCase(null)]
        [TestCase("")]
        [TestCase("   ")]
        [TestCase("verbose")]
        public void GivenMissingOrInvalidName_WhenParsing_ThenReturnsFallback(string input)
            => LogLevelParser.Parse(input, LogLevel.Warning).Should().Be(LogLevel.Warning);
    }

    [TestFixture]
    public class NullLoggerTests
    {
        [Test]
        public void NullLogger_IsNeverEnabled_AndDoesNotThrow()
        {
            NullLogger.Instance.IsEnabled(LogLevel.Error).Should().BeFalse();
            NullLogger.Instance.Error("x"); // must not throw
        }
    }

    [TestFixture]
    public class FileLoggerTests
    {
        private string _path;

        [SetUp]
        public void SetUp() =>
            _path = Path.Combine(Path.GetTempPath(), "t2s-logtest-" + Guid.NewGuid().ToString("N") + ".log");

        [TearDown]
        public void TearDown()
        {
            if (File.Exists(_path)) File.Delete(_path);
        }

        [Test]
        public void GivenMinLevelInfo_ThenDebugIsFilteredButInfoAndAboveAreWritten()
        {
            var log = new FileLogger(_path, LogLevel.Info);

            log.IsEnabled(LogLevel.Debug).Should().BeFalse();
            log.IsEnabled(LogLevel.Info).Should().BeTrue();

            log.Debug("dbg-should-not-appear");
            log.Info("info-should-appear");
            log.Warning("warn-should-appear");

            var text = File.ReadAllText(_path);
            text.Should().NotContain("dbg-should-not-appear");
            text.Should().Contain("info-should-appear");
            text.Should().Contain("warn-should-appear");
            text.Should().Contain("[INFO]");
        }

        [Test]
        public void GivenLevelNone_ThenNothingIsWritten()
        {
            var log = new FileLogger(_path, LogLevel.None);

            log.Error("nope");

            File.Exists(_path).Should().BeFalse();
        }

        [Test]
        public void GivenException_ThenItsDetailIsIncluded()
        {
            var log = new FileLogger(_path, LogLevel.Warning);

            log.Warning("boom", new InvalidOperationException("the-inner-message"));

            var text = File.ReadAllText(_path);
            text.Should().Contain("boom");
            text.Should().Contain("the-inner-message");
        }

        [Test]
        public void GivenTinyMaxBytes_ThenTheFileStaysBounded()
        {
            var log = new FileLogger(_path, LogLevel.Info, maxBytes: 200);

            for (int i = 0; i < 100; i++)
                log.Info("line " + i + " padding padding padding");

            new FileInfo(_path).Length.Should().BeLessThan(2000, "the log self-rotates past its cap");
        }
    }
}
