using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using ListMakerTwo;

namespace Tests.ListMakerTwo
{
    [TestFixture]
    public class SettingsHelperTests
    {
        [Test]
        public void SettingsHelperTests_CanSplitString()
        {
            var part1 = "";
            var part2 = "";
            SettingsHelper.ReadSpan("Part1:Part2", out part1, out part2);

            Assert.That(part1, Is.EqualTo("Part1"));
            Assert.That(part2, Is.EqualTo("Part2"));
        }

        [Test]
        public void SettingsHelperTests_CanSplitInt()
        {
            var part1 = 0;
            var part2 = 0;
            SettingsHelper.ReadSpan("111:222", out part1, out part2);

            Assert.That(part1, Is.EqualTo(111));
            Assert.That(part2, Is.EqualTo(222));
        }
    }
}
