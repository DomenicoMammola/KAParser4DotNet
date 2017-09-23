using System;
using Mammola.KAParser.mParser;

namespace mParserSampleApplication
{
  class Program
  {
    static void Main(string[] args)
    {
      mParser TempParser = new mParser();
      double TempValue = new Double();
      string TempValueStr = "";
      TempParser.Calculate("len('pippo')", ref TempValue);
      Console.WriteLine("len('pippo')=" + TempValue.ToString());            
      TempParser.Calculate("(1+1) + 1", ref TempValue);
      Console.WriteLine("(1+1) + 1=" + TempValue.ToString());
      TempParser.Calculate("2**2", ref TempValue);
      Console.WriteLine("pow(2,2)=" + TempValue.ToString());
      TempParser.Calculate("3*pi", ref TempValue);
      Console.WriteLine("3*pi=" + TempValue.ToString());
      
      TempParser.Calculate("_now_", ref TempValue);
      Console.WriteLine("_now_=" + TempValue.ToString());
      TempParser.Calculate("round(_now_,0)", ref TempValue);
      Console.WriteLine("round(_now_,0)=" + TempValue.ToString() + " " + DateTime.FromOADate(TempValue).ToString());
      TempParser.Calculate("today(1)", ref TempValue);
      Console.WriteLine("today(1)=" + TempValue.ToString());
      TempParser.Calculate("todouble('12.22')", ref TempValue);
      Console.WriteLine("todouble('12.22')=" + TempValue.ToString());
      TempParser.Calculate("todouble('12,22')", ref TempValue);
      Console.WriteLine("todouble('12,22')=" + TempValue.ToString());

      TempParser.Calculate("IF((2>1), 10, -10)", ref TempValue);
      Console.WriteLine("IF((2>1), 10, -10)=" + TempValue.ToString());

      TempParser.CalculateString("concatenate('paolino', 'paperino')", ref TempValueStr);
      Console.WriteLine("concatenate('paolino', 'paperino')=" + TempValueStr);
      TempParser.CalculateString("uppercase(concatenate('paolino', 'paperino'))", ref TempValueStr);
      Console.WriteLine("uppercase(concatenate('paolino', 'paperino'))=" + TempValueStr);
      Console.ReadLine();
    }
  }
}
