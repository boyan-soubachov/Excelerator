/* Project Name: ExcelIdea
 * Copyright (C) Boyan Soubachov, 2014
 * */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelIdea
{
    public enum Token_Type
    {
        Null,
        Constant,
        Cell,
        Range,
        Add,
        Subtract,
        Multiply,
        Divide,
        Exponent,
        LessThan,
        GreaterThan,
        Function,
        LeftParen,
        RightParen,
        Date,
        String,
        SheetReferenceStart,
        SheetReferenceEnd
    }

    class Token
    {
        public Token_Type Type;
        private object Value;
        public int Precedence;
        public string Operator_String;
        public int Function_Arguments = 0;

        public Token()
        {
            //null constructor
            Type = Token_Type.Null;
        }

        public Token(Token_Type Type, string Value)
        {
            this.Type = Type;
            this.Value = (string)Value;
            SetOperatorString(Type);
        }
        public Token(Token_Type Type, string Value, int precedence)
        {
            this.Type = Type;
            this.Value = (string)Value;
            this.Precedence = precedence;
            SetOperatorString(Type);
        }

        public Token(Token_Type Type, float Value)
        {
            this.Type = Type;
            this.Value = (float)Value;
            SetOperatorString(Type);
        }
        public Token(Token_Type Type, float Value, int precedence)
        {
            this.Type = Type;
            this.Value = (float)Value;
            this.Precedence = precedence;
            SetOperatorString(Type);
        }

        public string GetStringValue()
        {
            return (string)this.Value; //faster than ToString()
        }

        public float GetNumericValue()
        {
            return (float)this.Value;
        }

        public void SetValue(string value)
        {
            this.Value = value;
        }

        public void SetValue(float value)
        {
            this.Value = value;
        }

        private void SetOperatorString(Token_Type input)
        {
            switch (input)
            {
                case Token_Type.Null:
                    Operator_String = null;
                    break;
                case Token_Type.Constant:
                    Operator_String = "*-c";
                    break;
                case Token_Type.String:
                    Operator_String = "*-s";
                    break;
                case Token_Type.Cell:
                    Operator_String = "*-d";
                    break;
                case Token_Type.Add:
                    Operator_String = "+";
                    break;
                case Token_Type.Subtract:
                    Operator_String = "-";
                    break;
                case Token_Type.Multiply:
                    Operator_String = "*";
                    break;
                case Token_Type.Divide:
                    Operator_String = "/";
                    break;
                case Token_Type.Exponent:
                    Operator_String = "^";
                    break;
                case Token_Type.LessThan:
                    Operator_String = "<";
                    break;
                case Token_Type.GreaterThan:
                    Operator_String = ">";
                    break;
                case Token_Type.Function:
                    Operator_String = (string)this.Value;
                    break;
                case Token_Type.LeftParen:
                    Operator_String = "(";
                    break;
                case Token_Type.RightParen:
                    Operator_String = ")";
                    break;
                case Token_Type.Range:
                    Operator_String = ":";
                    break;
                case Token_Type.SheetReferenceStart:
                    Operator_String = "";
                    break;
                default:
                    Operator_String = "";
                    break;
            }
        }
    }
}
