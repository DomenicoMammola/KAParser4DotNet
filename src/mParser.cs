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
using System.Collections.Specialized;
using System.Collections.Generic;
using System.Globalization;
using Mammola.KAParser.mDateTimeUtilities;

namespace Mammola.KAParser.mParser
{

  public enum TmParserValueType 
  {
    vtFloat, 
    vtString
  }

  public delegate void ParserGetValueEventHandler(object sender, string valueName, ref double value, ref bool successful);
  public delegate void ParserGetStrValueEventHandler(object sender, string valueName, ref string value, ref bool successful);
  public delegate void ParserGetRangeValuesEventHandler(object sender, string func, TmParserValueType valueType, ref List<Object> valuesArray, ref bool successful);

  public delegate void ParserCalcUserFunctionEventHandler(object sender, string func, StringCollection parameters, ref double value, ref bool successful);
  public delegate void ParserCalcStrUserFunctionEventHandler(object sender, string func, StringCollection parameters, ref string value, ref bool handled);


  public sealed class mParser
  {
    const int sInvalidString = 1;
    const int sSyntaxError = 2;
    const int sFunctionError = 3;
    const int sWrongParamCount = 4;
    const int sFunctionUnknown = 5;

    const string mP_intfunc_trunc = "trunc";
    const string mP_intfunc_sin = "sin";
    const string mP_intfunc_cos = "cos";
    const string mP_intfunc_tan = "tan";
    const string mP_intfunc_frac = "frac";
    const string mP_intfunc_int = "int";

    const string mP_intconst_now = "_now_";
    const string mP_intconst_today = "_today_";
    const string mP_intconst_true = "true";
    const string mP_intconst_false = "false";
    const string mP_intconst_pi = "pi";

    const string mp_specfunc_if = "if";
    const string mp_specfunc_empty = "empty";
    const string mp_specfunc_len = "len";
    const string mp_specfunc_and = "and";
    const string mp_specfunc_or = "or";
    const string mp_specfunc_safediv = "safediv";
    const string mp_specfunc_concatenate = "concatenate";
    const string mp_specfunc_concat = "concat";
    const string mp_specfunc_repl = "repl";
    const string mp_specfunc_left = "left";
    const string mp_specfunc_right = "right";
    const string mp_specfunc_substr = "substr";
    const string mp_specfunc_tostr = "tostr";
    const string mp_specfunc_pos = "pos";
    const string mp_specfunc_uppercase = "uppercase";
    const string mp_specfunc_lowercase = "lowercase";
    const string mp_specfunc_compare = "compare";
    const string mp_specfunc_comparestr = "comparestr";
    const string mp_specfunc_comparetext = "comparetext";
    const string mp_specfunc_round = "round";
    const string mp_specfunc_ceil = "ceil";
    const string mp_specfunc_floor = "floor";
    const string mp_specfunc_not = "not";
    const string mp_specfunc_sum = "sum";
    const string mp_specfunc_max = "max";
    const string mp_specfunc_min = "min";
    const string mp_specfunc_avg = "avg";
    const string mp_specfunc_count = "count";
    const string mp_specfunc_getday = "getday";
    const string mp_specfunc_getweek = "getweek";
    const string mp_specfunc_getmonth = "getmonth";
    const string mp_specfunc_getyear = "getyear";
    const string mp_specfunc_startoftheweek = "startoftheweek";
    const string mp_specfunc_startofthemonth = "startofthemonth";
    const string mp_specfunc_endofthemonth = "endofthemonth";
    const string mp_specfunc_todate = "todate";
    const string mp_specfunc_todatetime = "todatetime";
    const string mp_specfunc_now = "now";
    const string mp_specfunc_today = "today";
    const string mp_specfunc_stringtodatetime = "stringtodatetime";
    const string mp_specfunc_todouble = "todouble";
    const string mp_specfunc_tonumber = "tonumber";
    const string mp_specfunc_between = "between";
    const string mp_specfunc_trim = "trim";
    const string mp_specfunc_ltrim = "ltrim";
    const string mp_specfunc_rtrim = "rtrim";

    const string mp_rangefunc_childsnotnull = "childsnotnull";
    const string mp_rangefunc_parentnotnull = "parentnotnull";
    const string mp_rangefunc_parentsnotnull = "parentsnotnull";
    const string mp_rangefunc_childs = "childs";
    const string mp_rangefunc_parents = "parents";
    const string mp_rangefunc_parent = "parent";

    private enum TCalculationType
    {
      calculateValue,
      calculateFunction
    }

    private enum TFormulaToken
    {
      tkUndefined,
      tkEOF, tkERROR,
      tkLBRACE, tkRBRACE, tkNUMBER, tkIDENT, tkSEMICOLON, // 7
      tkPOW, // 6
      tkINV, tkNOT, //5
      tkMUL, tkDIV, tkMOD, tkPER, //4
      tkADD, tkSUB, // 3
      tkLT, tkLE, tkEQ, tkNE, tkGE, tkGT, //2
      tkOR, tkXOR, tkAND, tkSTRING  //1
    }


    private class TLexResult
    {
      private bool FIsInteger;
      private bool FIsDouble;
      private bool FIsUndefined;
      public int IntValue;
      public double DoubleValue;
      public string StringValue;
      public TFormulaToken Token;

      public void Clear()
      {     
          FIsUndefined = true;
          Token = TFormulaToken.tkUndefined;
      }

      public bool IsInteger()
      { 
        return FIsInteger;    
      }

      public bool IsString()
      {
        return (! FIsDouble) & (! FIsInteger) & (! FIsUndefined);
      }

      public bool IsDouble()
      {
        return FIsDouble;
      }
    
      public void SetInteger (int aValue)
      {
        FIsDouble = false;
        FIsUndefined = false;
        FIsInteger = true;
        Token = TFormulaToken.tkNUMBER;
        IntValue = aValue;
      }

      public void SetDouble (double aValue)
      {
        FIsDouble = true;
        FIsUndefined = false;
        FIsInteger = false;
        Token = TFormulaToken.tkNUMBER;
        DoubleValue = aValue;
      }

      public void SetString (string aValue)
      {
        FIsDouble = false;
        FIsUndefined = false;
        FIsInteger = false;
        Token = TFormulaToken.tkSTRING;
        StringValue = aValue;
      }        
    }
  
  private class TLexState
  {
    private string FFormula;
    private int FCharIndex;
    private int FLenFormula;

    public bool IsEof()
    {
      return (FCharIndex >= FLenFormula); 
    }

    public char currentChar()
    { 
      return FFormula[FCharIndex];
    }

    public void Advance()
    { 
      FCharIndex++;
    }

    public void GoBack()
    { 
      FCharIndex--;
    }

    public string Formula
    { 
      get {return FFormula;}
      set 
       {
         FFormula = value;
         FLenFormula = FFormula.Length;
         FCharIndex = 0;
       }
    }
    
    public int CharIndex
    {
      get { return FCharIndex; }
    }
  }

       
    private bool IsDecimalSeparator (char aValue)
    { return (aValue == '.') || (aValue == ','); }

    private bool IsStringSeparator (char aValue)
    { return (aValue == '\'') || (aValue == '\"');}

    
    private string CleanFormula(string aFormula)
    {
	  if (aFormula == null) {
		  return "";
	  } else {
		  string temp =  aFormula.Trim();
		  temp = temp.Replace((char)10, '\0');
		  temp = temp.Replace((char)13, '\0');
		  return temp;
		  
	  }
    }

    private bool DoInternalCalculate(TCalculationType calculation, string subFormula, ref double resValue)
    {
      bool successful = false;

      if (calculation == TCalculationType.calculateValue)
      {
        if (String.Equals(subFormula, mP_intconst_now, StringComparison.OrdinalIgnoreCase))
        {
          resValue = DateTime.Now.ToOADate();
          successful = true;
        }
        else
        {
          if (String.Equals(subFormula, mP_intconst_today, StringComparison.OrdinalIgnoreCase))
          {
            resValue = DateTime.Today.ToOADate();
            successful = true;
          }
          else
          {
            if (String.Equals(subFormula, mP_intconst_true, StringComparison.OrdinalIgnoreCase))
            {
              resValue = 1;
              successful = true;
            }
            else
            {
              if (String.Equals(subFormula, mP_intconst_false, StringComparison.OrdinalIgnoreCase))
              {
                resValue = 0;
                successful = true;
              }
              else
              {
                if (String.Equals(subFormula, mP_intconst_pi, StringComparison.OrdinalIgnoreCase))
                {
                  resValue = Math.PI;
                  successful = true;
                }
                else
                {
                  if (this.OnGetValue != null)
                  { 
                    this.OnGetValue(this, subFormula, ref resValue, ref successful);
                  }
                  else
                  { successful = false; }
                }
              }
            }
          }
        }
      }
      else // calculation = calculateFunction
      {
        if (String.Equals(subFormula, mP_intfunc_trunc, StringComparison.OrdinalIgnoreCase))
        { 
          resValue = (int)Math.Truncate(resValue);
          successful = true;
        }
        else
        {
          if (String.Equals(subFormula, mP_intfunc_sin, StringComparison.OrdinalIgnoreCase))
          {
            resValue = Math.Sin(resValue);
            successful = true;
          }
          else
          {
            if (String.Equals(subFormula, mP_intfunc_cos, StringComparison.OrdinalIgnoreCase))
            {
              resValue = Math.Cos(resValue);
              successful = true;
            }
            else
            {
              if (String.Equals(subFormula, mP_intfunc_tan, StringComparison.OrdinalIgnoreCase))
              {
                resValue = Math.Tan(resValue);
                successful = true;
              }
              else
              { 
                if (String.Equals(subFormula, mP_intfunc_frac, StringComparison.OrdinalIgnoreCase))
                {
                  resValue = resValue - (int)resValue;
                  successful = true;
                }
                else
                {
                  if (String.Equals(subFormula, mP_intfunc_int, StringComparison.OrdinalIgnoreCase))
                  {
                    resValue = (int) resValue;
                    successful = true;
                  }
                  else
                  {
                    if (this.OnGetValue != null)
                    {
                      this.OnGetValue(this, subFormula, ref resValue, ref successful);
                    }
                    else
                    {
                      successful = false;
                    }
                  }
                }
              }
            }
          }
        }
      }
      return successful;
    }


    private bool DoInternalStringCalculate(TCalculationType calculation, string subFormula, ref string resValue)
    {
      bool successful = false;
      if (calculation == TCalculationType.calculateFunction)
      {
        if (String.Equals(subFormula, mp_specfunc_trim, StringComparison.OrdinalIgnoreCase))
        { 
          resValue = resValue.Trim(); 
          successful = true;
        }          
        else
        {
          if (String.Equals(subFormula, mp_specfunc_ltrim, StringComparison.OrdinalIgnoreCase))
          { 
            resValue = resValue.TrimStart (); 
            successful = true;
          }          
          else
          {
            if (String.Equals(subFormula, mp_specfunc_rtrim, StringComparison.OrdinalIgnoreCase))
            { 
              resValue = resValue.TrimEnd(); 
              successful = true;
            }
            else
            {
              if (this.OnGetStrValue != null)
              { 
                this.OnGetStrValue(this, subFormula, ref resValue, ref successful);
                
              }
              else
              { successful = false; }
            }
          }
         }     
       }
       return successful;
     }

    private bool DoInternalRangeCalculate(string RangeFunction, string Func, TmParserValueType ValueType, ref List<Object> ValuesArray)
    { 
      bool successful = false;

      if (this.OnGetRangeValues != null)
      {
        this.OnGetRangeValues(this, Func, ValueType, ref ValuesArray, ref successful);
      } 
      return successful;
    }
  
    private double DoUserFunction(string funct, ref StringCollection ParametersList)
    { 
      bool successful = false;
      double TempValue = new Double();

      if (this.OnCalcUserFunction != null)
      {
        this.OnCalcUserFunction(this, funct, ParametersList, ref TempValue, ref successful);
        if (!successful)
        {
          RaiseError(sFunctionError, funct);
          return 0;
        }
        else
        {
          return TempValue;
        }
      }
      else      
      { 
        RaiseError(sFunctionError, funct);
        return 0;
      }
   }

    private string DoStrUserFunction(string funct, ref StringCollection ParametersList)
    { 
      string TempValue = "";

      bool successful = false;
      if (this.OnCalcStrUserFunction != null)
      {
        this.OnCalcStrUserFunction(this, funct, ParametersList,ref TempValue, ref successful);
        if (! successful)
        {
          RaiseError(sFunctionError, funct);
          return "";
        }
        else
        {
          return TempValue;
        }
      }
      else
      {
        RaiseError(sFunctionError, funct);
        return "";
      }  
    }

    // Lexical Analyzer Function http://www.gnu.org/software/bison/manual/html_node/Lexical.html
    private void yylex (ref TLexState lexState, ref TLexResult lexResult)
    { 
      string currentIntegerStr = "";
      int currentInteger; 
      double currentFloat;
      double currentdecimal;
      string currentString = "";
      char currentCharSeparator;
      char opChar;
    
      lexResult.Clear();

      if (lexState.IsEof())
      {
        lexResult.Token = TFormulaToken.tkEOF;
        return;
      }
     
      while (lexState.currentChar().Equals(' ') && (! lexState.IsEof()))
      { lexState.Advance(); }        

      if (lexState.IsEof())
      {
        lexResult.Token = TFormulaToken.tkEOF;
        return;
      }
        
      // ------------ forse una stringa..

      if (IsStringSeparator(lexState.currentChar()))
      {
        currentString = "";
        currentCharSeparator = lexState.currentChar();

        lexState.Advance();

        bool doLoop = true;
        while (doLoop)
        {
          if (lexState.IsEof())
          {
            RaiseError(sInvalidString);
            return;
          }
          if (lexState.currentChar().Equals(currentCharSeparator))
          {
            doLoop = false;
          }
          else
          {
            currentString = currentString + lexState.currentChar().ToString();
            lexState.Advance();
          }                      
        }

        lexState.Advance();
        lexResult.SetString(currentString);
        return;
      }

      // ------------ forse un numero..

      if (Char.IsNumber(lexState.currentChar()))
      {
        currentIntegerStr = lexState.currentChar().ToString();
        lexState.Advance();
        
        while ((! lexState.IsEof()) && (Char.IsNumber(lexState.currentChar())))
        {
          currentIntegerStr = currentIntegerStr + lexState.currentChar().ToString();
          lexState.Advance();
         }

        if (! Int32.TryParse(currentIntegerStr, out currentInteger))
        {
          lexResult.Token = TFormulaToken.tkERROR;
          return;
        }

        if ((! lexState.IsEof()) && (IsDecimalSeparator(lexState.currentChar())))
        {
          lexState.Advance();
          currentFloat = currentInteger;

          currentdecimal = 1;
          while ((! lexState.IsEof()) && (Char.IsNumber(lexState.currentChar())))
          {
            currentdecimal = currentdecimal / 10;
            currentFloat = currentFloat + (currentdecimal * Int32.Parse(lexState.currentChar().ToString()));
            lexState.Advance();
          }

          lexResult.SetDouble(currentFloat);
        }
        else
        {
          lexResult.SetInteger(currentInteger);
        }

        return;

      }


      // ------------ forse un identificatore..

      if ( (Char.IsLetter(lexState.currentChar())) || (lexState.currentChar().Equals('_')) ||
        (lexState.currentChar().Equals('@')) )
      {
        currentString = lexState.currentChar().ToString();
        lexState.Advance();

        while ( (! lexState.IsEof()) && ( (Char.IsLetterOrDigit(lexState.currentChar())) || (lexState.currentChar().Equals('_')) ||
          (lexState.currentChar().Equals('@')) ))
        {
          currentString = currentString + lexState.currentChar().ToString();
          lexState.Advance();
        }

        lexResult.SetString(currentString);
        lexResult.Token = TFormulaToken.tkIDENT;
        return;
      }

      // ------------ forse un operatore..

      opChar = lexState.currentChar();
      lexState.Advance();      
      switch (opChar)
      {
        case '=':
          {
            if ((! lexState.IsEof()) && (lexState.currentChar().Equals('=')))
            {
              lexState.Advance();
              lexResult.Token = TFormulaToken.tkEQ;
            }
            else
            {
              lexResult.Token = TFormulaToken.tkERROR;
            }
            break;
          }
        case '+':
          {
            lexResult.Token = TFormulaToken.tkADD;
            break;
          }
        case '-':
          {
            lexResult.Token = TFormulaToken.tkSUB;
            break;
          }
        case '*':
          {
            if ((! lexState.IsEof()) && (lexState.currentChar().Equals('*')))
            {
              lexState.Advance();
              lexResult.Token = TFormulaToken.tkPOW;
            }
            else
            {
              lexResult.Token = TFormulaToken.tkMUL;
            }
            break;
          }
        case '/':
          {
            lexResult.Token = TFormulaToken.tkDIV;
            break;
          }
        case '%':
          {
            if ((! lexState.IsEof()) && (lexState.currentChar().Equals('%')))
            {
              lexState.Advance();
              lexResult.Token = TFormulaToken.tkPER;
            }
            else
            {
              lexResult.Token = TFormulaToken.tkMOD;
            }
            break;
          }
        case '~':
          {
            lexResult.Token = TFormulaToken.tkINV;
            break;
          }
        case '^':
          {
            lexResult.Token = TFormulaToken.tkXOR;
            break;
          }
        case '&':
          {
            lexResult.Token = TFormulaToken.tkAND;
            break;
          }
        case '|':
          {
            lexResult.Token = TFormulaToken.tkOR;
            break;
          }
        case '<':
          {
            if ((! lexState.IsEof()) && (lexState.currentChar().Equals('=')))
            {
              lexState.Advance();
              lexResult.Token = TFormulaToken.tkLE;
            }
            else
            {
              if ((! lexState.IsEof()) && (lexState.currentChar().Equals('>')))
              {
                lexState.Advance();
                lexResult.Token = TFormulaToken.tkNE;
              }
              else
              {
                lexResult.Token = TFormulaToken.tkLT;
              }
            }
            break;
          }
        case '>':
          {
            if ((! lexState.IsEof()) && (lexState.currentChar().Equals('=')))
            {
              lexState.Advance();
              lexResult.Token = TFormulaToken.tkGE;
            }
            else
            {
              if ((! lexState.IsEof()) && (lexState.currentChar().Equals('<')))
              {
                lexState.Advance();
                lexResult.Token = TFormulaToken.tkNE;
              }
              else
              {
                lexResult.Token = TFormulaToken.tkGT;
              }
            }
            break;
          }
        case '!':
          {
            if ((! lexState.IsEof()) && (lexState.currentChar().Equals('=')))
            {
              lexState.Advance();
              lexResult.Token = TFormulaToken.tkNE;
            }
            else
            {
              lexResult.Token = TFormulaToken.tkNOT;
            }
            break;
          }
        case '(':
          {
             lexResult.Token = TFormulaToken.tkLBRACE;
             break;
          }
        case ')':
          {
            lexResult.Token = TFormulaToken.tkRBRACE;
            break;
          }
        case ';':
          {
            lexResult.Token = TFormulaToken.tkSEMICOLON;
            break;
          }
        default:
        {
          lexResult.Token = TFormulaToken.tkERROR;
          lexState.GoBack();
          break;
        }
      }
    }

    // double
    private void  StartCalculate(ref double resValue, ref TLexState lexState, ref TLexResult lexResult)
    { 
      Calc6(ref resValue, ref lexState, ref lexResult);
      while (lexResult.Token.Equals(TFormulaToken.tkSEMICOLON))
      {
        yylex(ref lexState, ref lexResult);
        Calc6(ref resValue, ref lexState, ref lexResult);
      }
      if (! (lexResult.Token.Equals(TFormulaToken.tkEOF))) 
      {
        RaiseError(sSyntaxError);    
      }
    }

    private void Calc6(ref double resValue, ref TLexState lexState, ref TLexResult lexResult)
    {     
      double NewDouble = new Double();
      TFormulaToken LastToken;

      Calc5(ref resValue, ref lexState, ref lexResult);
      while (lexResult.Token.Equals(TFormulaToken.tkOR) ||
        lexResult.Token.Equals(TFormulaToken.tkXOR) ||
        lexResult.Token.Equals(TFormulaToken.tkAND))       
      {
        LastToken = lexResult.Token;
        yylex(ref lexState, ref lexResult);
        Calc5(ref NewDouble, ref lexState, ref lexResult);
        switch (LastToken)
        {
          case TFormulaToken.tkOR : 
            {
              resValue = (int)Math.Truncate(resValue) |  (int)Math.Truncate(NewDouble);
              break;
            }
          case TFormulaToken.tkAND: 
            {          
              resValue =(int) Math.Truncate(resValue) & (int)Math.Truncate(NewDouble);
              break;
            }
          case TFormulaToken.tkXOR: 
            {
              resValue = (int)Math.Truncate(resValue) ^ (int)Math.Truncate(NewDouble);
              break;
            }
        }
      }
    }

    private void Calc5(ref double resValue, ref TLexState lexState, ref TLexResult lexResult)
    { 
      double NewDouble = new Double();
      TFormulaToken  LastToken;
    
      Calc4(ref resValue, ref lexState, ref lexResult);
      while ((lexResult.Token.Equals(TFormulaToken.tkLT) ||
        lexResult.Token.Equals(TFormulaToken.tkLE) ||
        lexResult.Token.Equals(TFormulaToken.tkEQ) ||
        lexResult.Token.Equals(TFormulaToken.tkNE) ||
        lexResult.Token.Equals(TFormulaToken.tkGE) ||
        lexResult.Token.Equals(TFormulaToken.tkGT)))      
      {
        LastToken = lexResult.Token;
        yylex(ref lexState, ref lexResult);
        Calc4(ref NewDouble, ref lexState, ref lexResult);
        switch (LastToken)
        {
          case TFormulaToken.tkLT: 
            {
              resValue = BooleanToFloat((! internalDoublesAreEqual(resValue, NewDouble)) && (resValue < NewDouble));
              break;
            }
          case TFormulaToken.tkLE: 
            {
              resValue = BooleanToFloat(internalDoublesAreEqual(resValue, NewDouble) || (resValue < NewDouble));
              break;
            }
          case TFormulaToken.tkEQ: 
            {
              resValue = BooleanToFloat(internalDoublesAreEqual(resValue, NewDouble));
              break;
            }
          case TFormulaToken.tkNE: 
            {
              resValue = BooleanToFloat(! internalDoublesAreEqual(resValue, NewDouble));
              break;
            }
          case TFormulaToken.tkGE: 
            {
              resValue = BooleanToFloat(internalDoublesAreEqual(resValue, NewDouble) || (resValue > NewDouble));
              break;
            }
          case TFormulaToken.tkGT: 
            {
              resValue = BooleanToFloat((! internalDoublesAreEqual(resValue, NewDouble)) && (resValue > NewDouble));
              break;
            }
        }
      }    
    }

    private void  Calc4(ref double resValue, ref TLexState lexState, ref TLexResult lexResult)
    {     
      double NewDouble = new Double();
      TFormulaToken LastToken;
    
      Calc3(ref resValue, ref lexState, ref lexResult);
      while (
        (lexResult.Token.Equals(TFormulaToken.tkADD)) ||
        (lexResult.Token.Equals(TFormulaToken.tkSUB)))
      {
        LastToken = lexResult.Token;
        yylex(ref lexState, ref lexResult);
        Calc3(ref NewDouble, ref lexState, ref lexResult);
        switch (LastToken)
        {
          case TFormulaToken.tkADD: 
            {
              resValue = resValue + NewDouble;
              break;
            }
          case TFormulaToken.tkSUB: 
            {
              resValue = resValue - NewDouble;
              break;
            }
        }
      }
    }

    private void  Calc3(ref double resValue, ref TLexState lexState, ref TLexResult lexResult)
    {    
      double NewDouble = new Double();
      TFormulaToken LastToken;
    
      Calc2(ref resValue, ref lexState, ref lexResult);
      while (
        lexResult.Token.Equals(TFormulaToken.tkMUL) ||
        lexResult.Token.Equals(TFormulaToken.tkDIV) ||
        lexResult.Token.Equals(TFormulaToken.tkMOD) ||
        lexResult.Token.Equals(TFormulaToken.tkPER) )
      {
        LastToken = lexResult.Token;
        yylex(ref lexState, ref lexResult);
        Calc2(ref NewDouble, ref lexState, ref lexResult);
        switch (LastToken)
        {
          case TFormulaToken.tkMUL: 
            {
              resValue = resValue * NewDouble;
              break;
            }
          case TFormulaToken.tkDIV: 
            { 
              resValue = resValue / NewDouble;
              break;
            }
          case TFormulaToken.tkMOD: // resto della divisione tra interi
            { 
              resValue = (int)(Math.Truncate(resValue)) % (int)(Math.Truncate(NewDouble));
              break;
            }
          case TFormulaToken.tkPER: 
            { 
              resValue = resValue * NewDouble / 100;
              break;
            }
        }
      }
    }

    private void Calc2(ref double resValue, ref TLexState lexState, ref TLexResult lexResult)
    {     
      TFormulaToken LastToken;
    
      if (lexResult.Token.Equals(TFormulaToken.tkNOT) ||
        lexResult.Token.Equals(TFormulaToken.tkINV) ||
        lexResult.Token.Equals(TFormulaToken.tkADD) ||
        lexResult.Token.Equals(TFormulaToken.tkSUB))
      {
        LastToken = lexResult.Token;
        yylex(ref lexState, ref lexResult);
        Calc2(ref resValue, ref lexState, ref lexResult);
        switch (LastToken)
        {
          case TFormulaToken.tkNOT:
            {
              if (Math.Truncate(resValue) == 0)
              {
                resValue = 1;
              }
              else
              {
                resValue = 0;
              }
              break;
            }
          case TFormulaToken.tkINV: 
          { 
            resValue = (~ (int)(Math.Truncate(resValue)));
            break;
          }
          case TFormulaToken.tkADD: 
          {
             resValue = +(resValue);
             break;
          };
          case TFormulaToken.tkSUB: 
          {
            resValue = -(resValue);
            break;
          }
        }
      }
      else
      {
        Calc1(ref resValue, ref lexState, ref lexResult);    
      }
    
    }

    private void Calc1(ref double resValue, ref TLexState lexState, ref TLexResult lexResult)
    { 
      double NewDouble = new Double();

      CalcTerm(ref resValue, ref lexState, ref lexResult);
      if (lexResult.Token.Equals(TFormulaToken.tkPOW))
      {
        yylex(ref lexState, ref lexResult);
        CalcTerm(ref NewDouble, ref lexState, ref lexResult);
        resValue = Math.Pow(resValue, NewDouble);
      }   
    }

    private void CalcTerm(ref double resValue, ref TLexState lexState, ref TLexResult lexResult)
    { 
      string currentIdent;
      StringCollection ParamList = new StringCollection();
    
      switch (lexResult.Token)
      {
        case TFormulaToken.tkNUMBER:
        {
          if (lexResult.IsInteger())
          {
            resValue = lexResult.IntValue;
          }
          else
          {
            resValue = lexResult.DoubleValue;
          }
          yylex(ref lexState, ref lexResult);
          break;
        }
        case TFormulaToken.tkLBRACE:
        {
          yylex(ref lexState, ref lexResult);
          Calc6(ref resValue, ref lexState, ref lexResult);
          if (lexResult.Token.Equals(TFormulaToken.tkRBRACE))
          {
            yylex(ref lexState, ref lexResult);
          }
          else
          {
            RaiseError(sSyntaxError);
          }
          break;
        }
        case TFormulaToken.tkIDENT:
        {
          currentIdent = lexResult.StringValue;
          yylex(ref lexState, ref lexResult);
          if (lexResult.Token.Equals(TFormulaToken.tkLBRACE))
          {
            if (IsFunction(currentIdent))
            {
              ParamList.Clear();
              ParseFunctionParameters(currentIdent, ref ParamList, ref lexState);
              resValue = ExecuteFunction(currentIdent, ref ParamList);
              yylex(ref lexState, ref lexResult);
            }
            else
            if (IsInternalFunction(currentIdent))
            {
              yylex(ref lexState, ref lexResult);
              Calc6(ref resValue, ref lexState, ref lexResult);
              if (lexResult.Token.Equals(TFormulaToken.tkRBRACE))
              {
                yylex(ref lexState, ref lexResult);
              }
              else
              {
                RaiseError(sFunctionError, currentIdent);
              }
              if (! DoInternalCalculate(TCalculationType.calculateFunction, currentIdent, ref resValue))
              {
                RaiseError(sFunctionError, currentIdent);
              }
            }
            else
            {
              ParamList.Clear();
              ParseFunctionParameters(currentIdent, ref ParamList, ref lexState);
              resValue = DoUserFunction(currentIdent, ref ParamList);
              
              yylex(ref lexState, ref lexResult);
            }
          }
          else
          {
            if (! DoInternalCalculate(TCalculationType.calculateValue, currentIdent, ref resValue))
            {
              RaiseError(sFunctionError, currentIdent);
            }
          }
          break;
        }
        default:
        {
          RaiseError(sSyntaxError);
          break;
        }
      }
    
    }

    // string
    private void StartCalculateStr(ref string resValue, ref TLexState lexState, ref TLexResult lexResult)
    { 
      CalculateStrLevel1(ref resValue, ref lexState, ref lexResult);
      while (lexResult.Token.Equals(TFormulaToken.tkSEMICOLON))
      {
        yylex(ref lexState, ref lexResult);
        CalculateStrLevel1(ref resValue, ref lexState, ref lexResult);
      }
      if (! (lexResult.Token.Equals(TFormulaToken.tkEOF)))
      {
        RaiseError(sSyntaxError);    
      }
    }

    private void CalculateStrLevel1(ref string resValue, ref TLexState lexState, ref TLexResult lexResult)
    {     
      string newString = "";
    
      CalculateStrLevel2(ref resValue, ref lexState, ref lexResult);
      while (lexResult.Token.Equals(TFormulaToken.tkADD))
      {
        yylex(ref lexState, ref lexResult);
        CalculateStrLevel2(ref newString, ref lexState, ref lexResult);
        resValue = resValue + newString;
      }
    }

    private void CalculateStrLevel2(ref string resValue, ref TLexState lexState, ref TLexResult lexResult)
    {     
      string currentIdent = "";
      StringCollection ParamList = new StringCollection();
    
      switch (lexResult.Token)
      {
        case TFormulaToken.tkSTRING:
        {
          resValue = lexResult.StringValue.Substring(0, lexResult.StringValue.Length);
          yylex(ref lexState, ref lexResult);
          break;
        }
        case TFormulaToken.tkLBRACE:
        {
          yylex(ref lexState, ref lexResult);
          CalculateStrLevel1(ref resValue, ref lexState, ref lexResult);
          if (lexResult.Token.Equals(TFormulaToken.tkRBRACE))
          {
            yylex(ref lexState, ref lexResult);
          }
          else
          {
            RaiseError(sSyntaxError);
          }
          break;
        }
        case TFormulaToken.tkIDENT:
        {
          currentIdent = lexResult.StringValue;
          yylex(ref lexState, ref lexResult);
          if (lexResult.Token.Equals(TFormulaToken.tkLBRACE))
          {
            ParamList.Clear();

            ParseFunctionParameters(currentIdent, ref ParamList, ref lexState);

            if (IsFunction(currentIdent))
            {
              resValue = ExecuteStrFunction(currentIdent, ref ParamList);
            }
            else
            {
              if (ParamList.Count == 1)
              {
                CalculateString(ParamList[0], ref resValue);
                if (! DoInternalStringCalculate(TCalculationType.calculateFunction, currentIdent, ref resValue))
                {
                  resValue = DoStrUserFunction(currentIdent, ref ParamList);
                }
              }
              else
              {
                resValue = DoStrUserFunction(currentIdent, ref ParamList);
              }
            }
            yylex(ref lexState, ref lexResult);
          }
          else
          {
            if (! DoInternalStringCalculate(TCalculationType.calculateValue, currentIdent, ref resValue))
            {
              RaiseError(sFunctionError, currentIdent);
            }
          }
          break;
        }
        default:
        {
          RaiseError(sSyntaxError);
          break;
        }
      }
    
    }
    // range functions

    private string InsideBraces (string aValue)
    {     
      aValue = aValue.Trim();      
      
      int len = aValue.Length;

      if (len <= 2)
      {
        RaiseError(sSyntaxError);  // stringa vuota o '()'
        return "";
      }

      if (! aValue[0].Equals('('))
      {
        RaiseError(sSyntaxError);  // non inizia con '('
        return "";
      }

      if (! aValue[len - 1].Equals(')'))
      {
        RaiseError(sSyntaxError);  // non finisce con ')'
        return "";
      }

      return aValue.Substring(1, len -2).Trim();      
    }

    private void ManageRangeFunction (string funct, ref StringCollection ParametersList, TmParserValueType ValueType, ref List<Object> TempValuesArray)
    {  
      double TempDouble;
      string TempString;
      int k;
    
      TempValuesArray.Clear();

      if (ParametersList.Count == 1)
      {
        string lowercaseFunct = funct.ToLower();
        
        k = lowercaseFunct.IndexOf(mp_rangefunc_childsnotnull);
        if (k >= 0) 
        {
          TempString = funct.Substring(k + mp_rangefunc_childsnotnull.Length - 1);
          DoInternalRangeCalculate(mp_rangefunc_childsnotnull, InsideBraces(TempString), ValueType, ref TempValuesArray);
          return;
        }
        k = lowercaseFunct.IndexOf(mp_rangefunc_parentnotnull);
        if (k >= 0) 
        {
          TempString = funct.Substring(k + mp_rangefunc_parentnotnull.Length - 1);
          DoInternalRangeCalculate(mp_rangefunc_parentnotnull, InsideBraces(TempString), ValueType, ref TempValuesArray);
          return;
        }
        k = lowercaseFunct.IndexOf(mp_rangefunc_parentsnotnull);
        if (k >= 0) 
        {
          TempString = funct.Substring(k + mp_rangefunc_parentsnotnull.Length - 1);
          DoInternalRangeCalculate(mp_rangefunc_parentsnotnull, InsideBraces(TempString), ValueType, ref TempValuesArray);
          return;
        }
        k = lowercaseFunct.IndexOf(mp_rangefunc_childs);
        if (k >= 0) 
        {
          TempString = funct.Substring(k + mp_rangefunc_childs.Length - 1);
          DoInternalRangeCalculate(mp_rangefunc_childs, InsideBraces(TempString), ValueType, ref TempValuesArray);
          return;
        }
        k = lowercaseFunct.IndexOf(mp_rangefunc_parent);
        if (k >= 0) 
        {
          TempString = funct.Substring(k + mp_rangefunc_parent.Length - 1);
          DoInternalRangeCalculate(mp_rangefunc_parent, InsideBraces(TempString), ValueType, ref TempValuesArray);
          return;
        }
        k = lowercaseFunct.IndexOf(mp_rangefunc_parents);
        if (k >= 0) 
        {
          TempString = funct.Substring(k + mp_rangefunc_parents.Length - 1);
          DoInternalRangeCalculate(mp_rangefunc_parents, InsideBraces(TempString), ValueType, ref TempValuesArray);
          return;
        }
      }

      for (int i = 0; i < ParametersList.Count; i++)
      {
        if (ValueType == TmParserValueType.vtFloat)
        {
          TempDouble = new Double();
          Calculate(ParametersList[i], ref TempDouble);
          TempValuesArray.Add(TempDouble);          
        }
        else
        {
          TempString = "";
          CalculateString(ParametersList[i], ref TempString);
          TempValuesArray.Add(TempString);
        }
      }
     }


    private bool IsInternalFunction(string functionName)
    { 
      string temp = functionName.ToLowerInvariant();
      
      return  ((temp==mP_intfunc_trunc) || (temp==mP_intfunc_sin) || (temp==mP_intfunc_cos) || (temp==mP_intfunc_tan) ||
        (temp==mP_intfunc_frac) || (temp==mP_intfunc_int));
    }


    private bool IsFunction(string functionName)
    { 
    
      string temp = functionName.ToLowerInvariant();
      return ((temp == mp_specfunc_if) ||
        (temp == mp_specfunc_empty) ||
        (temp == mp_specfunc_len) ||
        (temp == mp_specfunc_and) ||
        (temp == mp_specfunc_or) ||
        (temp == mp_specfunc_safediv) ||
        (temp == mp_specfunc_concatenate) ||
        (temp == mp_specfunc_concat) ||
        (temp == mp_specfunc_repl) ||
        (temp == mp_specfunc_left) ||
        (temp == mp_specfunc_right) ||
        (temp == mp_specfunc_substr) ||
        (temp == mp_specfunc_tostr) ||
        (temp == mp_specfunc_pos) ||
        (temp == mp_specfunc_uppercase) ||
        (temp == mp_specfunc_lowercase) ||
        (temp == mp_specfunc_compare) ||
        (temp == mp_specfunc_comparestr) ||
        (temp == mp_specfunc_comparetext) ||
        (temp == mp_specfunc_round) ||
        (temp == mp_specfunc_ceil) ||
        (temp == mp_specfunc_floor) ||
        (temp == mp_specfunc_not) ||
        (temp == mp_specfunc_sum) ||
        (temp == mp_specfunc_max) ||
        (temp == mp_specfunc_min) ||
        (temp == mp_specfunc_avg) ||
        (temp == mp_specfunc_count) ||
        (temp == mp_specfunc_getday) ||
        (temp == mp_specfunc_getweek) ||
        (temp == mp_specfunc_getmonth) ||
        (temp == mp_specfunc_getyear) ||
        (temp == mp_specfunc_startoftheweek) ||
        (temp == mp_specfunc_startofthemonth) ||
        (temp == mp_specfunc_endofthemonth) ||
        (temp == mp_specfunc_todate) ||
        (temp == mp_specfunc_todatetime) ||
        (temp == mp_specfunc_now) ||
        (temp == mp_specfunc_today) ||
        (temp == mp_specfunc_stringtodatetime) ||
        (temp == mp_specfunc_todouble) ||
        (temp == mp_specfunc_tonumber));
    
    }

    private void ParseFunctionParameters(string funct, ref StringCollection ParametersList, ref TLexState lexState)
    {     
      int startindex = lexState.CharIndex;
      int i;
      int q;
      int par = 1;
      bool insideCommas = false;
      string parametersStr;

      // questa funzione deve cercare la prima parentesi ( e poi andare avanti verso dx
      // aprendo e chiudendo le eventuali parentesi
      // fino a trovare la parentesi ) che corrisponde alla prima trovata
      // cosi' si risolve il problema delle parentesi annidate

      while (! lexState.IsEof())
      {
        if (lexState.Formula[lexState.CharIndex].Equals('\''))
        {
          insideCommas = ! insideCommas;
        }
        else
        if ((!insideCommas) && (lexState.Formula[lexState.CharIndex].Equals('(')))
        {
          par++;
        }
        else        
        if ((!insideCommas) && (lexState.Formula[lexState.CharIndex].Equals(')')))
        {
          par--;
        }
        

        if (par == 0)
        {
          break;
        }
        else
        {
          lexState.Advance();
        }
      }

      parametersStr = lexState.Formula.Substring(startindex, (lexState.CharIndex - startindex)).Trim();      

      if (par != 0)
      {
        RaiseError(sFunctionError, funct);
      }
      if (parametersStr == "")
      {
        return; // no parameters
      }

      lexState.Advance(); // spostiamoci al carattere successivo alla parentesi

      par = 0;
      i = 0;
      int k = 0;
      q = 0;
      insideCommas = false;

      while (i < parametersStr.Length)
      {
        if (parametersStr[i].Equals('\''))
        {
          insideCommas = ! insideCommas;
        }
        if ((! insideCommas) && (parametersStr[i].Equals ('(')))
        {
          par++;
        }
        else
        if ((! insideCommas) && (parametersStr[i].Equals(')')))
        {
          par--;
        }
        if ((! insideCommas) && parametersStr[i].Equals(',') && (par == 0))
        {          
          ParametersList.Add(parametersStr.Substring(q, k).Trim());
          q = i + 1;
          k = 0;
        }
        else
        {
          k++;
        }
        i++;
        
      }
      
      ParametersList.Add(parametersStr.Substring(q).Trim());
    
    }

    private double ExecuteFunction(string funct, ref StringCollection ParametersList)
    {     
      string TempStrValue;
      string TempStrValue2;
      double TempDouble;
      double TempDouble2;
      double TempDouble3;
      double TempDouble4;
      double TempDouble5;
      double TempDouble6;
      bool TempBoolean;
      DateTime TempDateTime;            
      double TempResult;
      
      if (String.Equals(funct, mp_specfunc_if,StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 3)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempResult = new Double();
        Calculate(ParametersList[0], ref TempResult);
        if (TempResult != 0)
        {
          Calculate(ParametersList[1], ref TempResult);
        }
        else
        {
          Calculate(ParametersList[2], ref TempResult);
        }
        return TempResult;
      }
      else
      if (String.Equals(funct, mp_specfunc_len, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 1)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempStrValue = "";
        CalculateString(ParametersList[0], ref TempStrValue);
        TempResult = TempStrValue.Length;
        return TempResult;
      }
      else
      if (String.Equals(funct, mp_specfunc_pos, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 2)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempStrValue = "";
        TempStrValue2 = "";
        CalculateString(ParametersList[0], ref TempStrValue);
        CalculateString(ParametersList[1], ref TempStrValue2);
        TempResult = TempStrValue2.IndexOf(TempStrValue);
        return TempResult;
      }      
      else      
      if (String.Equals(funct, mp_specfunc_todouble, StringComparison.OrdinalIgnoreCase) ||
         (String.Equals(funct, mp_specfunc_tonumber, StringComparison.OrdinalIgnoreCase)))
      {
        if (ParametersList.Count != 1)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempStrValue = "";
        CalculateString(ParametersList[0], ref TempStrValue);
        TempResult =  StrToFloatExt(TempStrValue);
        return TempResult;
      }
      else
      if (String.Equals(funct, mp_specfunc_today, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 1)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempDouble = new Double();
        Calculate(ParametersList[0], ref TempDouble);
        return (DateTime.Today.Date.ToOADate() - TempDouble);
      }
      else
      if (String.Equals(funct, mp_specfunc_now, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 1)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempDouble = new Double();
        Calculate(ParametersList[0], ref TempDouble);
        return (DateTime.Today.ToOADate() - TempDouble);
      }
      else
      if (String.Equals(funct, mp_specfunc_getday, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 1)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempDouble = new Double();
        Calculate(ParametersList[0], ref TempDouble);
        TempDateTime = (DateTime.FromOADate(TempDouble));        
        return DateTime.DaysInMonth(TempDateTime.Year, TempDateTime.Month);
      }
      else
      if (String.Equals(funct, mp_specfunc_getyear, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 1)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempDouble = new Double();
        Calculate(ParametersList[0], ref TempDouble);
        TempDateTime = (DateTime.FromOADate(TempDouble));
        return TempDateTime.Year;
      }
      else
      if (String.Equals(funct, mp_specfunc_getweek, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 1)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempDouble = new Double();
        Calculate(ParametersList[0], ref TempDouble);
        TempDateTime = (DateTime.FromOADate(TempDouble));
        return DateTimeUtilities.WeekOfYear_ISO8601(TempDateTime);
      }
      else
      if (String.Equals(funct, mp_specfunc_getmonth, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 1)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempDouble = new Double();
        Calculate(ParametersList[0], ref TempDouble);
        TempDateTime = (DateTime.FromOADate(TempDouble));
        return TempDateTime.Month;        
      }
      else
      if (String.Equals(funct, mp_specfunc_startoftheweek, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 1)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempDouble = new Double();
        Calculate(ParametersList[0], ref TempDouble);
        TempDateTime = (DateTime.FromOADate(TempDouble));
        return DateTimeUtilities.StartOfWeek(TempDateTime).ToOADate();        
      }
      else
      if (String.Equals(funct, mp_specfunc_startofthemonth, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 1)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempDouble = new Double();
        Calculate(ParametersList[0], ref TempDouble);
        TempDateTime = (DateTime.FromOADate(TempDouble));
        return DateTimeUtilities.FirstDayOfMonth(TempDateTime).ToOADate();        
      }
      else
      if (String.Equals(funct, mp_specfunc_endofthemonth, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 1)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempDouble = new Double();
        Calculate(ParametersList[0], ref TempDouble);
        TempDateTime = (DateTime.FromOADate(TempDouble));
        return DateTimeUtilities.LastDayOfMonth(TempDateTime).ToOADate();        
      }
      else
      if ((String.Equals(funct, mp_specfunc_todate, StringComparison.OrdinalIgnoreCase)))
      {
        if (ParametersList.Count != 3)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempDouble = new Double();
        TempDouble2 = new Double();
        TempDouble3 = new Double();
        Calculate(ParametersList[0], ref TempDouble);
        Calculate(ParametersList[1], ref TempDouble2);
        Calculate(ParametersList[2], ref TempDouble3);
        TempDateTime = (new DateTime((int)Math.Round(TempDouble), (int)Math.Round(TempDouble2), (int)Math.Round(TempDouble3)));
        return TempDateTime.ToOADate();
      }
      else
      if (String.Equals(funct, mp_specfunc_todatetime, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 6)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempDouble = new Double();
        TempDouble2 = new Double();
        TempDouble3 = new Double();
        TempDouble4 = new Double();
        TempDouble5 = new Double();
        TempDouble6 = new Double();
        Calculate(ParametersList[0], ref TempDouble);
        Calculate(ParametersList[1], ref TempDouble2);
        Calculate(ParametersList[2], ref TempDouble3);
        Calculate(ParametersList[3], ref TempDouble4);
        Calculate(ParametersList[4], ref TempDouble5);
        Calculate(ParametersList[5], ref TempDouble6);
        TempDateTime = (new DateTime((int)Math.Round(TempDouble), (int)Math.Round(TempDouble2), (int)Math.Round(TempDouble3),
          (int)Math.Round(TempDouble4), (int)Math.Round(TempDouble5), (int)Math.Round(TempDouble6)));        
        return TempDateTime.ToOADate();
      }
      else
      if (String.Equals(funct, mp_specfunc_empty, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 1)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempStrValue = "";
        CalculateString(ParametersList[0], ref TempStrValue);
        return BooleanToFloat(TempStrValue.Trim().Length == 0);
      }
      else
      if (String.Equals(funct, mp_specfunc_compare, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 2)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempStrValue = "";
        TempStrValue2 = "";
        CalculateString(ParametersList[0], ref TempStrValue);
        CalculateString(ParametersList[1], ref TempStrValue2);
        return BooleanToFloat( (String.Compare(TempStrValue, TempStrValue2) == 0));
      }
      else
      if (String.Equals(funct, mp_specfunc_comparestr, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 2)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempStrValue = "";
        TempStrValue2 = "";
        CalculateString(ParametersList[0], ref TempStrValue);
        CalculateString(ParametersList[1], ref TempStrValue2);
        return (String.Compare(TempStrValue, TempStrValue2, false));        
      }
      else
      if (String.Equals(funct, mp_specfunc_comparetext, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 2)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempStrValue = "";
        TempStrValue2 = "";        
        CalculateString(ParametersList[0], ref TempStrValue);
        CalculateString(ParametersList[1], ref TempStrValue2);
        return (String.Compare(TempStrValue, TempStrValue2, true));        
      }
      else
      if (String.Equals(funct, mp_specfunc_and, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count < 2)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempDouble = new Double();
        Calculate(ParametersList[0], ref TempDouble);
        TempBoolean = FloatToBoolean(TempDouble);
        if (TempBoolean)
        {
          for (int i = 1; i <= (ParametersList.Count - 1);  i++)
          {
            TempDouble = new Double();
            Calculate(ParametersList[i], ref TempDouble);
            TempBoolean = TempBoolean && FloatToBoolean(TempDouble);
            if (! TempBoolean)
              { break; }              
          }
        }
        return BooleanToFloat(TempBoolean);
      }
      else
      if (String.Equals(funct, mp_specfunc_or, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count < 2)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempDouble = new Double();
        Calculate(ParametersList[0], ref TempDouble);
        TempBoolean = FloatToBoolean(TempDouble);
        if (! TempBoolean)
        {
          for (int i = 1; i <= (ParametersList.Count - 1);  i++)
          {
            TempDouble = new Double();
            Calculate(ParametersList[i], ref TempDouble);
            TempBoolean = TempBoolean || FloatToBoolean(TempDouble);
            if (TempBoolean)
              { break; }              
          }
        }
        return BooleanToFloat(TempBoolean);
      }
      else
      if (String.Equals(funct, mp_specfunc_safediv, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 2)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempDouble = new Double();
        TempDouble2 = new Double();
        Calculate(ParametersList[0], ref TempDouble);
        Calculate(ParametersList[1], ref TempDouble2);
        
        return mFloatsManagement.mFloatsManager.Instance.SafeDiv(TempDouble, TempDouble2);
      }
      else
      if (String.Equals(funct, mp_specfunc_between, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 3)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempDouble = new Double();
        TempDouble2 = new Double();
        TempDouble3 = new Double();
        Calculate(ParametersList[0], ref TempDouble);
        Calculate(ParametersList[1], ref TempDouble2);
        Calculate(ParametersList[2], ref TempDouble3);

        return BooleanToFloat(
          internalDoublesAreEqual(TempDouble, TempDouble2) ||
          internalDoublesAreEqual(TempDouble, TempDouble3) ||
          ((TempDouble >= TempDouble2) && (TempDouble <= TempDouble3)));
      }
      else
      if (String.Equals(funct, mp_specfunc_round, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 2)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempDouble = new Double();
        TempDouble2 = new Double();
        Calculate(ParametersList[0], ref TempDouble);
        Calculate(ParametersList[1], ref TempDouble2);

        return Math.Round(TempDouble, (int)Math.Round(TempDouble2)); // vanno invertiti??       
      }
      else
      if (String.Equals(funct, mp_specfunc_ceil, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 1)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempDouble = new Double();
        Calculate(ParametersList[0], ref TempDouble);
        return Math.Ceiling(TempDouble);        
      }
      else
      if (String.Equals(funct, mp_specfunc_floor, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 1)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempDouble = new Double();
        Calculate(ParametersList[0], ref TempDouble);
        return Math.Floor(TempDouble);        
      }
      else
      if (String.Equals(funct, mp_specfunc_stringtodatetime, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 2)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempStrValue = "";
        TempStrValue2 = "";
        CalculateString(ParametersList[0], ref TempStrValue);
        CalculateString(ParametersList[1], ref TempStrValue2);
        // forse vanno invertite!!
        return DateTime.ParseExact(TempStrValue2, TempStrValue, System.Globalization.CultureInfo.InvariantCulture).ToOADate();                                  
      }
      else
      if (String.Equals(funct, mp_specfunc_not, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 1)
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        TempDouble = new Double();
        Calculate(ParametersList[0], ref TempDouble);
        if (FloatToBoolean(TempDouble))
        {
          return 0;
        }
        else
        {
          return 1;
        }
      }
      else
      if (String.Equals(funct, mp_specfunc_sum, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count < 1) 
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        List<Object> TempValuesArray = new List<Object>();

        ManageRangeFunction(funct, ref ParametersList, TmParserValueType.vtFloat, ref TempValuesArray);

        TempDouble = new Double();
        TempDouble = 0;
        for (int i = 0; i < TempValuesArray.Count; i++)
        {
          TempDouble = TempDouble + (double)TempValuesArray[i];
        }
        return TempDouble;
      }
      else
      if (String.Equals(funct, mp_specfunc_max, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count < 1) 
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        List<Object> TempValuesArray = new List<Object>();

        ManageRangeFunction(funct, ref ParametersList, TmParserValueType.vtFloat, ref TempValuesArray);
        
        TempDouble = new Double();

        if (TempValuesArray.Count == 0)
        {
          TempDouble = 0;
        }
        else
        {
          TempDouble = (double)TempValuesArray[0];
        }
        for (int i = 1; i < TempValuesArray.Count; i++)
        {
          if ((double)TempValuesArray[i] > TempDouble) 
          {
            TempDouble = (double)TempValuesArray[i];
          }          
        }
        return TempDouble;
      }
      else
      if (String.Equals(funct, mp_specfunc_min, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count < 1) 
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        List<Object> TempValuesArray = new List<Object>();

        ManageRangeFunction(funct, ref ParametersList, TmParserValueType.vtFloat, ref TempValuesArray);

        TempDouble = new Double();

        if (TempValuesArray.Count == 0)
        {
          TempDouble = 0;
        }
        else
        {
          TempDouble = (double)TempValuesArray[0];
        }
        for (int i = 1; i < TempValuesArray.Count; i++)
        {
          if ((double)TempValuesArray[i] < TempDouble) 
          {
            TempDouble = (double)TempValuesArray[i];
          }          
        }
        return TempDouble;
      }
      else
      if (String.Equals(funct, mp_specfunc_avg, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count < 1) 
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        List<Object> TempValuesArray = new List<Object>();

        ManageRangeFunction(funct, ref ParametersList, TmParserValueType.vtFloat, ref TempValuesArray);

        TempDouble = new Double();

        TempDouble = 0;
        for (int i = 0; i < TempValuesArray.Count; i++)
        {
          TempDouble = TempDouble + (double)TempValuesArray[i];
        }
        if (TempValuesArray.Count > 0)
        {
          TempDouble = TempDouble / TempValuesArray.Count;
        }
        return TempDouble;
      }
      else
      if (String.Equals(funct, mp_specfunc_count, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count < 1) 
        {
          RaiseError(sWrongParamCount);
          return 0;
        }
        List<Object> TempValuesArray = new List<Object>();

        ManageRangeFunction(funct, ref ParametersList, TmParserValueType.vtFloat, ref TempValuesArray);
        return TempValuesArray.Count;
      }
      else
      {
        RaiseError(sFunctionUnknown);
        return 0;
      }
    }


    private string ExecuteStrFunction(string funct, ref StringCollection ParametersList)
    {     
      string TempStrValue;
      string TempStrValue2;
      double TempValue = new Double();
      double TempValue2 = new Double();      


      if (String.Equals(funct, mp_specfunc_if, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 3)
        {
          RaiseError(sWrongParamCount);
          return "";
        }

        Calculate(ParametersList[0], ref TempValue);

        if (internalDoublesAreEqual(1, TempValue) || (TempValue > 1))
        {
          TempStrValue = "";
          CalculateString(ParametersList[1], ref TempStrValue);
        }
        else
        {
          TempStrValue = "";
          CalculateString(ParametersList[2], ref TempStrValue);
        }

        return TempStrValue;
      }
      else
      if (String.Equals(funct, mp_specfunc_repl, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 2)
        {
          RaiseError(sWrongParamCount);
          return "";
        }
        TempStrValue = "";
        CalculateString(ParametersList[0], ref TempStrValue);
        TempStrValue2 = TempStrValue;
		Calculate(ParametersList[1], ref TempValue);
        for (int i = 1; i <= (int)Math.Round(TempValue); i++)
        {
          TempStrValue2 = TempStrValue2 + TempStrValue;
        }
        return TempStrValue2;
      }
      else
      if (String.Equals(funct, mp_specfunc_concatenate, StringComparison.OrdinalIgnoreCase) ||
        String.Equals(funct, mp_specfunc_concat, StringComparison.OrdinalIgnoreCase)) 
      {
        if (ParametersList.Count < 2) 
        {
          RaiseError(sWrongParamCount);
          return "";
        }

        TempStrValue = "";
        TempStrValue2 = "";
        for (int i = 0; i < ParametersList.Count; i++)
        {
          CalculateString(ParametersList[i], ref TempStrValue);
          TempStrValue2 = TempStrValue2 + TempStrValue;
        }
        return TempStrValue2;
      }
      else
      if (String.Equals(funct, mp_specfunc_tostr, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 1)
        {
          RaiseError(sWrongParamCount);
          return "";
        }
        Calculate(ParametersList[0], ref TempValue);
        return TempValue.ToString();
      }
      else
      if (String.Equals(funct, mp_specfunc_uppercase, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 1) 
        {
          RaiseError(sWrongParamCount);
          return "";
        }
        TempStrValue = "";
        CalculateString(ParametersList[0], ref TempStrValue);
        return TempStrValue.ToUpper();
      }
      else
      if (String.Equals(funct, mp_specfunc_lowercase, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 1)
        {
          RaiseError(sWrongParamCount);
          return "";
        }
        TempStrValue = "";
        CalculateString(ParametersList[0], ref TempStrValue);
        return TempStrValue.ToLower();
      }
      else
      if (String.Equals(funct, mp_specfunc_left, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 2) 
        {
          RaiseError(sWrongParamCount);
          return "";
        }
        TempStrValue = "";
        CalculateString(ParametersList[0], ref TempStrValue);
        Calculate(ParametersList[1], ref TempValue);

        return TempStrValue.Substring(0, Math.Min(TempStrValue.Length, (int)Math.Round(TempValue)));        
      }
      else
      if (String.Equals(funct, mp_specfunc_right, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 2) 
        {
          RaiseError(sWrongParamCount);
          return "";
        }
        TempStrValue = "";
        CalculateString(ParametersList[0], ref TempStrValue);
        Calculate(ParametersList[1], ref TempValue);
        return TempStrValue.Substring(Math.Max(0,TempStrValue.Length - (int)Math.Round(TempValue)));        
      }
      else
      if (String.Equals(funct, mp_specfunc_substr, StringComparison.OrdinalIgnoreCase))
      {
        if (ParametersList.Count != 3)
        {
          RaiseError(sWrongParamCount);
          return "";
        }
        TempStrValue = "";
        TempStrValue2 = "";
        CalculateString(ParametersList[0], ref TempStrValue);
        Calculate(ParametersList[1], ref TempValue);
        Calculate(ParametersList[2], ref TempValue2);
        return TempStrValue.Substring((int)Math.Round(TempValue) - 1, (int)Math.Round(TempValue2) - 1); // -1, alla pascal!
      }
      else
      if (String.Equals(funct, mp_specfunc_sum, StringComparison.OrdinalIgnoreCase))
      {
        List<Object> TempValuesArray = new List<Object>();
        ManageRangeFunction(funct, ref ParametersList, TmParserValueType.vtString, ref TempValuesArray);        
        TempStrValue = "";
        for (int i = 0; i < TempValuesArray.Count; i++)
        {
          TempStrValue = TempStrValue + (string)TempValuesArray[i];
        }
        return TempStrValue;
      }
      else
      if (String.Equals(funct, mp_specfunc_max, StringComparison.OrdinalIgnoreCase))
      {
        List<Object> TempValuesArray = new List<Object>();
        ManageRangeFunction(funct, ref ParametersList, TmParserValueType.vtString, ref TempValuesArray);
        TempStrValue = "";
        for (int i = 0; i < TempValuesArray.Count; i++)
        {
          if (TempStrValue.CompareTo((string)TempValuesArray[i]) < 0)
          {
            TempStrValue = (string)TempValuesArray[i];
          }          
        }
        return TempStrValue;
      }
      else
      if (String.Equals(funct, mp_specfunc_min, StringComparison.OrdinalIgnoreCase))
      {
        List<Object> TempValuesArray = new List<Object>();
        ManageRangeFunction(funct, ref ParametersList, TmParserValueType.vtString, ref TempValuesArray);
        TempStrValue = "";
        for (int i = 0; i < TempValuesArray.Count; i++)
        {
          if (TempStrValue.CompareTo((string)TempValuesArray[i]) > 0)
          {
            TempStrValue = (string)TempValuesArray[i];
          }          
        }
        return TempStrValue;
      }
      else
      if (String.Equals(funct, mp_specfunc_count, StringComparison.OrdinalIgnoreCase))
      {
        List<Object> TempValuesArray = new List<Object>();
        ManageRangeFunction(funct, ref ParametersList, TmParserValueType.vtString, ref TempValuesArray);
        return TempValuesArray.Count.ToString();
      }
      else
      {
        RaiseError(sFunctionUnknown);
        return "";
      }

    }

    private void RaiseError (int aErrorCode)
    { 
      throw new System.Exception(aErrorCode.ToString()); 
    }

    private void RaiseError (int aErrorCode, string aErrorMessage)
    { 
      throw new System.Exception(aErrorCode.ToString()  + ' ' + aErrorMessage);       
    }
    
    private double BooleanToFloat (bool aValue)
    {
      if (aValue)
      { return 1; }
      else
      { return 0; };
    }

    private bool FloatToBoolean (double aValue)
    { return (Math.Round(Math.Abs(aValue)) >= 1); }

    private double StrToFloatExt (string aValue)
    {
      double returnValue = 0;

      // float.Parse(aValue, CultureInfo.InvariantCulture.NumberFormat);

      if (aValue.IndexOf(',') >= 0)
      { 
        aValue = aValue.Replace(',','.');
      }
      if (double.TryParse(aValue, System.Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, out returnValue))
      { 
        return returnValue;
      }
      else
      if (double.TryParse(aValue, System.Globalization.NumberStyles.Any, CultureInfo.CurrentCulture, out returnValue))
      { 
        return returnValue;
      }
      else
      if (double.TryParse(aValue, System.Globalization.NumberStyles.Any, CultureInfo.GetCultureInfo("en-US"), out returnValue))
      {
        return returnValue;
      }
      else 
      {
        throw new System.ArgumentException(aValue + " is not a valid number");
      }
    }


    private bool internalDoublesAreEqual (double aValue1, double aValue2)
    { 
      if (DecimalNumbers >= 0)
      { return mFloatsManagement.mFloatsManager.Instance.DoublesAreEqual(aValue1, aValue2, DecimalNumbers);  }
      else
      { return mFloatsManagement.mFloatsManager.Instance.DoublesAreEqual(aValue1, aValue2); }  
    }
  
    public mParser()
    { }

/*    ~mParser()
    { }*/


    public bool Calculate(string formula, ref double resValue)
    { 
      string newFormula = CleanFormula(formula);
      TLexState LexState = new TLexState();
      TLexResult LexResult = new TLexResult();

      LexState.Formula = newFormula;

      yylex(ref LexState, ref LexResult);

      StartCalculate(ref resValue, ref LexState, ref LexResult);

      return true;
    }

    public bool CalculateString(string formula, ref string resValue)
    { 
      string newFormula = CleanFormula(formula);
      TLexState LexState = new TLexState();
      TLexResult LexResult = new TLexResult();

      LexState.Formula = newFormula;

      yylex(ref LexState, ref LexResult);

      StartCalculateStr(ref resValue, ref LexState, ref LexResult);

      return true;
    }
     
    public int DecimalNumbers;   // se -1, prende il default in mFloatsManagement

    // per estrarre i valori ad esempio da un dataset
    public ParserGetValueEventHandler OnGetValue;
    public ParserGetStrValueEventHandler OnGetStrValue;
    public ParserGetRangeValuesEventHandler OnGetRangeValues;
    // per implementare funzioni custom
    public ParserCalcUserFunctionEventHandler OnCalcUserFunction;
    public ParserCalcStrUserFunctionEventHandler OnCalcStrUserFunction;

  }

}
