using Microsoft.VisualStudio.TestTools.UnitTesting;
using Mammola.KAParser.mParser;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mammola.KAParser.mParser.Tests
{
  [TestClass()]
  public class mParserTests
  {
    [TestMethod()]
    public void mParserTest()
    {
      Assert.Inconclusive();
    }

    [TestMethod()]
    public void CalculateTest()
    {
      mParser TempParser = new mParser();
      double TempValue = new Double();      
      TempParser.Calculate("(1+1) + 1", ref TempValue);
      Assert.AreEqual(TempValue, 3);      
      TempParser.Calculate("2**2", ref TempValue);
      Assert.AreEqual(TempValue, 4);
      TempParser.Calculate("4-(2*2)", ref TempValue);
      Assert.AreEqual(TempValue, 0);
      TempParser.Calculate("( 4 -(  2* 2  ))", ref TempValue);
      Assert.AreEqual(TempValue, 0);      
      TempParser.Calculate("(_today_)", ref TempValue);
      Assert.IsTrue(DateTime.FromOADate(TempValue).Equals(DateTime.Today));
      TempParser.Calculate("TODOUBLE('12')", ref TempValue);
      Assert.AreEqual(TempValue, 12);
      TempParser.Calculate("TODOUBLE('12,22')", ref TempValue);
      Assert.AreEqual(TempValue, 12.22);
      TempParser.Calculate("TODOUBLE('12.22')", ref TempValue);
      Assert.AreEqual(TempValue, 12.22);
      TempParser.Calculate("TODOUBLE(concatenate('12.', '22'))", ref TempValue);
      Assert.AreEqual(TempValue, 12.22);
      TempParser.Calculate("IF((2>1), 10, -10)", ref TempValue);
      Assert.AreEqual(TempValue, 10);
      TempParser.Calculate("comparetext('pippo', 'PIPPO')", ref TempValue);
      Assert.AreEqual(0, TempValue);

    }

    [TestMethod()]
    public void CalculateStringTest()
    {
      mParser TempParser = new mParser();      
      string TempValueStr = "";
      TempParser.CalculateString("concatenate('paolino', 'paperino')", ref TempValueStr);
      Assert.AreEqual("paolinopaperino", TempValueStr);      
      TempParser.CalculateString("uppercase(concatenate('paolino', 'paperino'))", ref TempValueStr);
      Assert.AreEqual("PAOLINOPAPERINO", TempValueStr);
      TempParser.CalculateString("if(len(concatenate('paolino', 'paperino')) > 100, 'maggiore di 100', 'minore di 100')", ref TempValueStr);
      Assert.AreEqual("minore di 100", TempValueStr);

    }
  }
}