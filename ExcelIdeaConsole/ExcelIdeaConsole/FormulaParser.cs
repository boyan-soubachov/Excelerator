/* Project Name: ExcelIdea
 * Copyright (C) Boyan Soubachov, 2014
 * */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.Diagnostics;

namespace ExcelIdea
{
    class FormulaParser
    {
        private static readonly HashSet<string> hs_operators = new HashSet<string>(new string[10] { "+", "-", "*", "/", "^", "%", "!", ":", "<", ">" });
        private static readonly HashSet<string> hs_la_operators = new HashSet<string>(new string[7] { "+", "-", "*", "/", ":", "<", ">" });
        
        public List<Token> Parse_Tokens(string formula)
        {
            List<Token> output = new List<Token>();
            Stack<Token> op = new Stack<Token>();
            Stack<int> arg_count = new Stack<int>();
            StringBuilder str_temp = new StringBuilder(formula);
            int i = 0;
            str_temp.EnsureCapacity(16);

            //formula conditioning
            if (CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator != ".")
                str_temp.Replace(CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator, ".");
            if (CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator != ",")
                str_temp.Replace(',', ';');
            if (formula.Contains("_xlfn."))
                str_temp.Replace("_xlfn.", string.Empty);
            if (formula.Contains("$"))
                str_temp.Replace("$", string.Empty);

            str_temp.Replace(" ", string.Empty);
            if (formula[0] == '=')
                str_temp.Remove(0, 1);
            while (formula[0] == '+') //Remove leading pluses for Lotus 123 compatibility
                str_temp.Remove(0, 1);

            formula = str_temp.ToString();

            //Shunting yard algorithm implementation
            while (formula.Length > 0)
            {
                if (char.IsDigit(formula[0])) //if number
                {
                    str_temp.Clear();
                    i = 0;
                    while (char.IsDigit(formula[i]) || formula[i] == '.')
                    {
                        if (formula[i] == '.')
                            if (str_temp.ToString().Contains(".")) //malformed real number, throw error
                            {
                                //error handling here
                            }
                        str_temp.Append(formula[i]);
                        i++;
                        if (formula.Length == i) break;
                    }
                    formula = formula.Remove(0, i);
                    output.Add(new Token(Token_Type.Constant, float.Parse(str_temp.ToString()), 0));
                }
                else if (char.IsLetter(formula[0])) //if function or cell reference
                {
                    str_temp.Clear();
                    i = 0;
                    while (char.IsLetterOrDigit(formula[i]) || formula[i] == '.')
                    {
                        str_temp.Append(formula[i]);
                        i++;
                        if (formula.Length == i)
                            break;
                    }
                    formula = formula.Remove(0, i);

                    for (int j = 0; j < str_temp.Length; j++) //convert to upper case, faster than creating a new stringbuilder
                        str_temp[j] = char.ToUpper(str_temp[j]);

                    if (formula.Length == 0 && Tools.IsCell(str_temp.ToString())) //if a cell reference
                    {
                        output.Add(new Token(Token_Type.Cell, str_temp.ToString(), 0));
                    }
                    else if (formula.Length > 0)
                    {
                        if (formula[0] == '(' || str_temp.ToString().Contains(".")) //if function
                        {
                            op.Push(new Token(Token_Type.Function, str_temp.ToString(), 0));
                        }
                        else if (formula[0] != '!' && char.IsDigit(str_temp[str_temp.Length - 1])) //cell reference again
                        {
                            output.Add(new Token(Token_Type.Cell, str_temp.ToString(), 0));
                        }
                        else //must be a string
                        {
                            if (formula[0] != '!')
                                output.Add(new Token(Token_Type.String, str_temp.ToString()));
                            else
                                output.Add(new Token(Token_Type.SheetReferenceStart, str_temp.ToString()));
                        }
                    }
                }
                else if (formula[0] == ';') //if argument separator
                {
                    Token tmp = op.Peek();
                    arg_count.Push(arg_count.Pop() + 1);
                    while (tmp.Type != Token_Type.LeftParen && op.Count > 0)
                    {
                        op.Pop();
                        output.Add(tmp);
                        if (op.Count > 0)
                            tmp = op.Peek();
                        //else
                        //incomplete brackets error
                    }
                    formula = formula.Remove(0, 1);
                }
                else if (hs_operators.Contains(formula[0].ToString())) //if operator
                {
                    Token tmp = new Token();
                    switch (formula[0])
                    {
                        case '+':
                            tmp = new Token(Token_Type.Add, 0, 2);
                            break;
                        case '-':
                            tmp = new Token(Token_Type.Subtract, 0, 2);
                            break;
                        case '*':
                            tmp = new Token(Token_Type.Multiply, 0, 3);
                            break;
                        case '/':
                            tmp = new Token(Token_Type.Divide, 0, 3);
                            break;
                        case '^':
                            tmp = new Token(Token_Type.Exponent, 0, 4);
                            break;
                        case '%':
                            break;
                        case ':':
                            tmp = new Token(Token_Type.Range, 0, 10);
                            break;
                        case '!':
                            tmp = new Token(Token_Type.SheetReferenceEnd, 0, 9);
                            break;
                        default:
                            break;
                    }

                    if (tmp.Type == Token_Type.Null)
                    {
                        //unsupported token error
                        break;
                    }
                    formula = formula.Remove(0, 1);

                    if (op.Count > 0)
                    {
                        Token tmp2 = op.Peek();
                        while (op.Count > 0 && (hs_operators.Contains(tmp2.Operator_String) && (hs_la_operators.Contains(tmp.Operator_String) && 
                            tmp.Precedence <= tmp2.Precedence) || tmp.Precedence < tmp2.Precedence))
                        {
                            output.Add(op.Pop());
                            if (op.Count > 0)
                                tmp2 = op.Peek();
                        }
                    }
                    op.Push(tmp);
                }
                else if (formula[0] == '(') //if left parenthesis
                {
                    op.Push(new Token(Token_Type.LeftParen, 0, 0));
                    if (formula[1] != ')')
                        arg_count.Push(1);
                    else
                        arg_count.Push(0);

                    formula = formula.Remove(0, 1);
                }
                else if (formula[0] == ')') //right parenthesis
                {
                    Token tmp = op.Peek();
                    while (tmp.Type != Token_Type.LeftParen && op.Count > 0)
                    {
                        op.Pop();
                        output.Add(tmp);
                        tmp = op.Peek();
                    }

                    if (tmp.Type == Token_Type.LeftParen && op.Count > 0)
                        op.Pop();
                    else
                    {
                        //left parenthesis missing error
                        break;
                    }
                    if (op.Count > 0)
                    {
                        tmp = op.Peek();
                        if (tmp.Type == Token_Type.Function)
                        {
                            tmp.Function_Arguments = arg_count.Pop();
                            output.Add(tmp);
                            op.Pop();
                        }
                    }
                    formula = formula.Remove(0, 1);
                }
                else if (formula[0] == '\"') //string
                {
                    formula = formula.Remove(0, 1);
                    if (!formula.Contains('\"'))
                    {
                        //error here for missing string closing quotation
                        break;
                    }
                    str_temp = new StringBuilder(formula.Substring(0, formula.IndexOf('\"')));
                    formula = formula.Substring(formula.IndexOf('\"') + 1);
                    output.Add(new Token(Token_Type.String, str_temp.ToString()));
                }
                else if (formula[0] == '\'') //sheet reference start
                {
                    str_temp.Clear();
                    i = 1;
                    while (formula[i] != '\'')
                    {
                        str_temp.Append(formula[i]);
                        i++;
                        if (formula.Length == i)
                            break;
                    }
                    formula = formula.Remove(0, i + 1);
                    output.Add(new Token(Token_Type.SheetReferenceStart, str_temp.ToString().ToUpper()));
                }
                else
                {
                    throw new SystemException("Unhandled parser character!");
                }
            }

            while (op.Count > 0)
            {
                Token tmp = op.Pop();
                if (tmp.Type == Token_Type.LeftParen || tmp.Type == Token_Type.RightParen)
                {
                    //show error here for mismatched parenthesis
                }
                else
                    output.Add(tmp);
            }

            return output;
        }
    }
}
