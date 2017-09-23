/*
 *  This is part of the KAParser4DotNet Library
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 *
 * This software is distributed without any warranty.
 *
 * @author Domenico Mammola (mimmo71@gmail.com - www.mammola.net)
 *
 */
using System;

namespace Mammola.KAParser.mFloatsManagement
{
  public sealed class mFloatsManager 
  {
    private static mFloatsManager instance;
    private int DefaultDecimalNumbers = 5;
    private double DefaultCompareValue = 1;

    private mFloatsManager() {
      SetDefaultDecimalNumbers(5);
    }

    public static mFloatsManager Instance {
      get 
      {
        if (instance == null)
          { instance = new mFloatsManager(); }
        return instance;
      }
    }

    public bool DoublesAreEqual(double aValue1, double aValue2, int aDecimalNumbers) {
      double CompareValue = Math.Pow(10, (-1 * aDecimalNumbers)) - 
        Math.Pow(10, (-1 * (aDecimalNumbers + 1))) - 
        Math.Pow(10, (-1 * (aDecimalNumbers + 2)));  
      return  (Math.Abs(aValue1 - aValue2) <= CompareValue);
     }

    public bool DoublesAreEqual(double aValue1, double aValue2) {
      return (Math.Abs(aValue1 - aValue2) <= DefaultCompareValue);
     }

    public double SafeDiv (double numer, double denom) {
      if (denom == 0)
        {return 0; }    
      else
        {return numer / denom; }    
     }

    public void SetDefaultDecimalNumbers (int aDecimalNumbers) {
      DefaultCompareValue = Math.Pow(10, -1 * aDecimalNumbers) - Math.Pow(10, -1 * (aDecimalNumbers + 1)) - Math.Pow(10, -1 * (aDecimalNumbers + 2));
      DefaultDecimalNumbers = aDecimalNumbers;
     }

  }

}
