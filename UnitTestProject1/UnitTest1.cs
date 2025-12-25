using Microsoft.VisualStudio.TestTools.UnitTesting;
using PhoneBillCalculator;
using System;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {
        // Тест 1: Корректные данные
        [TestMethod]
        public void CalculateTariff_NormalInput_CalculatesCorrectly()
        {
            var (cost, extra) = MainWindow.CalculateTariff(150, 100, 0.3, 1.6);
            Assert.AreEqual(110.0, cost, 0.001); // 100*0.3 + 50*1.6 = 110
            Assert.AreEqual(50, extra);
        }

        // Тест 2: Очень большое число
        [TestMethod]
        public void CalculateTariff_MaxIntValue_NoOverflow()
        {
            var (cost, extra) = MainWindow.CalculateTariff(int.MaxValue, 200, 0.7, 1.6);
            Assert.IsTrue(cost > 0);
            Assert.AreEqual(int.MaxValue - 200, extra);
        }

        // Тест 3: Отрицательные числа
        [TestMethod]
        public void CalculateTariff_NegativeMinutes_ReturnsZero()
        {
            var (cost, extra) = MainWindow.CalculateTariff(-50, 200, 0.7, 1.6);
            Assert.AreEqual(0, cost, 0.001); 
            Assert.AreEqual(0, extra); 
        }

        // Тест 4: Граничный случай - ноль минут
        [TestMethod]
        public void CalculateTariff_ZeroMinutes_ReturnsZero()
        {
            var (cost, extra) = MainWindow.CalculateTariff(0, 200, 0.7, 1.6);
            Assert.AreEqual(0, cost, 0.001);
            Assert.AreEqual(0, extra);
        }
    }
}
