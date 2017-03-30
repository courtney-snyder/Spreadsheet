/*
 * Courtney Snyder
 * CptS 321, Homework 7
 * Last Updated: 3/30/2017
 * Description: This is where the logic for the spreadsheet app is at.
 */

using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq; //Contains XDocument contents

namespace SpreadsheetEngine
{
    public abstract class Cell
    {
        public int mRowIndex;
        public int mColumnIndex;
        protected string mText;
        protected string mValue;
        protected uint mBGColor;
        public ExpTree mTree;
        public event PropertyChangedEventHandler PropertyChanged = delegate { };

        public Cell() //Default constructor
        {
            this.mColumnIndex = 0;
            this.mRowIndex = 0;
            this.mText = "";
            this.mValue = "";
            this.mTree = new ExpTree("");
            this.mBGColor = 4294967295; //Start with a white background (0 is black)
        }

        public int ColumnIndex
        {
            get { return this.ColumnIndex; }
        }

        public int RowIndex
        {
            get { return this.mRowIndex; }
        }

        public string Text
        {
            set
            {
                if (mText == value) //If new value is the same as what was already stored in mText, don't make any changes
                    return;
                mText = value; //Update mText
                PropertyChanged(this, new PropertyChangedEventArgs("Text")); //Notify subscribers that Text update has been made
            }
            get { return this.mText; }
        }
        public string Value
        {
            get
            {
                return this.mValue;
            }
        }
        public uint BGColor
        {
            set
            {
                if (mBGColor == value) //If the color entered is not a uint or it is the same color as it was before
                    return;
                mBGColor = value; //Update mBGColor
                PropertyChanged(this, new PropertyChangedEventArgs("BGColor")); //Notify subscribers that BGColor has been made
            }
            get
            {
                return this.mBGColor;
            }
        }
    }

    public class Spreadsheet
    {
        private class SheetCell : Cell //The implementation of a Cell for Spreadsheet
        {
            public string Value
            {
                set
                {
                    this.mValue = value;
                }
            }
        }
        private int mColumnCount;
        private int mRowCount;
        private Stack<ICmd> mUndoStack;
        private Stack<ICmd> mRedoStack;
        //private Stack<MultiCmd> mUndoStack;
        //private Stack<MultiCmd> mRedoStack;
        private Cell[,] mSheet = new SheetCell [50,26];
        public Dictionary<string, HashSet<string>> mDependencies;
        public event PropertyChangedEventHandler CellPropertyChanged = delegate { };
        public Spreadsheet (int maxRows, int maxCols)
        {
            this.mColumnCount = maxCols;
            this.mRowCount = maxRows;
            this.mDependencies = new Dictionary<string, HashSet<string>>();
            int i = 0, j = 0;
            for (; i < maxRows; i++) //Separate loops for columns and rows in case they are not equal
            {
                for (j = 0; j < maxCols; j++)
                {
                    Cell temp = new SheetCell();
                    mSheet[i, j] = temp;
                    mSheet[i, j].mRowIndex = i; //Give that cell the correct row index
                    mSheet[i, j].mColumnIndex = j; //Give that cell the correct column index
                    mSheet[i, j].PropertyChanged += OnPropertyChanged; //Subscribe to that cell's PropertyChanged event
                }
            }
            mUndoStack = new Stack<ICmd>();
            mRedoStack = new Stack<ICmd>();
        }
        public Cell GetCell (int row, int col)
        {
            return mSheet[row,col];
        }
        public int ColumnCount
        {
            get { return mColumnCount; }
        }
        public int RowCount
        {
            get { return mRowCount; }
        }
        protected void OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            Cell temp = sender as Cell;
            if (e.PropertyName == "Text")
            {
                mSheet[temp.mRowIndex, temp.mColumnIndex] = temp;
                Evaluate(temp); //Evaluate the cell
            }
            if (e.PropertyName == "BGColor")
            {
                mSheet[temp.mRowIndex, temp.mColumnIndex] = temp;
                UpdateBGColor(temp); //Apply the color to the cell
            }
        }
        public void Evaluate(Cell toEvaluate) //Decide if the cell is a formula or not and updates value accordingly
        {
            SheetCell temp = toEvaluate as SheetCell;
            temp.Value = toEvaluate.Text; //Assume the cell isn't a formula; set mValue to mText
            temp.Text = toEvaluate.Text; //Set mText to mText
            toEvaluate = temp;
            if (temp.Text != "")
            {
                if (temp.Text.ElementAt(0) == '=') //If the value starts with =, the input is a formula and it must be evaluated
                {
                    string formula = temp.Text, sub = "";
                    int colIndex = 0, rowIndex = 0, i = 0;
                    Cell inFormula;
                    bool badref = false;
                    alphaBool alpha = alphaBool.NOT;
                    string cellName = toAlpha(toEvaluate.mColumnIndex.ToString()) + (toEvaluate.mRowIndex + 1).ToString();
                    formula = formula.Replace(" ", ""); //Get rid of any whitespace in the formula
                    formula = formula.Remove(0, 1); //Remove '='
                    temp.mTree = new ExpTree(formula);
                    Dictionary<string, double> variables = temp.mTree.GetVars(); //Retrieve the variables from the tree
                    var keys = variables.Keys; //Get the variable names
                    double varVal = 0;
                    for (i = 0; i < keys.Count; i++) //Lookup each variable in the spreadsheet
                    {
                        string thing = keys.ElementAt(i);
                        varVal = 0;
                        sub = thing.Substring(1); //Get the row index by separating the letter and number
                        alpha = isAlpha(thing.ElementAt(0));
                        if (alpha != alphaBool.NOT && int.TryParse(sub, out rowIndex) && rowIndex > 0 && rowIndex < 50) //If the variable has the form "A thru Z""1-50"
                        {
                            colIndex = (int)thing[0] - 'A'; //Get column index as an int
                            rowIndex--; //Got the rowIndex as an int in the conditional, now subtract one since rows go 1-50 and index goes 0-49
                            inFormula = GetCell(rowIndex, colIndex); //Get the cell in the formula (eg =A1, this would get A1)
                            double.TryParse(mSheet[rowIndex, colIndex].Value, out varVal);
                            temp.mTree.SetVar(thing, varVal); //Set the variable (if tryParse is unsuccessful, varVal = 0)
                        }
                        else //If the variable has a variable name that isn't a cell name
                        {
                            badref = true;
                            temp.Value = "!(bad reference)"; //Updated the value, not the text so user can hopefully figure out why it's a bad reference
                        }
                    }
                    if (!badref)
                    {
                        double result = temp.mTree.Eval();
                        temp.Value = result.ToString(); //Update current cell's value
                        UpdateTableAdd(temp); //Add the appropriate dependencies
                        UpdateTableDelete(temp); //Remove old dependencies
                        if (!checkSelfReference(cellName) && !checkCircularReference(cellName))
                        {
                            UpdateDependencies(temp); //Update current dependencies
                        }
                        else
                            temp.Value = "!(self reference)";
                    }
                    CellPropertyChanged(toEvaluate, new PropertyChangedEventArgs("Value")); //Update the UI
                }
            }

            else
            {
                toEvaluate.Text = "";
                CellPropertyChanged(toEvaluate, new PropertyChangedEventArgs("Value")); //Update the UI
            }
        }

        public void UpdateDependencies(Cell updatedCell)
        {
            string cellName = toAlpha(updatedCell.mColumnIndex) + (updatedCell.mRowIndex + 1).ToString(), sub = "";
            int rowIndex = 0, colIndex = 0;
            double cellValue = 0;
            double.TryParse(updatedCell.Value, out cellValue);
            if (this.mDependencies.ContainsKey(cellName)) //If the cell has dependencies
            {
                for (int i = 0; i < mDependencies[cellName].Count; i++)
                {
                    string current = mDependencies[cellName].ElementAt(i);
                    colIndex = (int)current[0] - 'A'; //Get column index as an int
                    sub = current.Substring(1); //Get the row index by separating the letter and number
                    rowIndex = int.Parse(sub); //Get row index as an int
                    rowIndex--;
                    mSheet[rowIndex, colIndex].mTree.SetVar(cellName, cellValue); //Update the current cell's value for cellName
                    Evaluate(mSheet[rowIndex, colIndex]);
                }
            }
        }

        public void UpdateTableDelete(Cell toUpdate)
        {
            string cellName = toAlpha(toUpdate.mColumnIndex) + (toUpdate.mRowIndex + 1).ToString();
            Dictionary<string, double> variables = toUpdate.mTree.GetVars(); //Retrieve the variable (cell) names from the tree
            var keys = variables.Keys;
            var dict = mDependencies.Keys;
            string thing = "";
            for (int i = 0; i < dict.Count; i++) //Look at each key in the dependency dictionary
            {
                thing = dict.ElementAt(i);
                if (keys.Count != 0) //If the current cell uses other variables, check that
                {
                    if (dict.Contains(thing) && !keys.Contains(thing)) //If the dependencies dictionary contains thing and that variable is not used in the current formula, remove the dependency
                    {
                        mDependencies[thing].Remove(cellName); //Remove the current cell from the dependency hash
                    }
                }
                else  //The current cell uses no variables, so If the dependencies dictionary contains thing, remove the dependency
                {
                    if (dict.Contains(thing))
                    {
                        mDependencies[thing].Remove(cellName); //Remove the current cell from the dependency hash
                    }
                }
            }
        }

        public void UpdateTableAdd(Cell toUpdate)
        {
            string cellName = toAlpha(toUpdate.mColumnIndex) + (toUpdate.mRowIndex+1).ToString();
            Dictionary<string, double> variables = toUpdate.mTree.GetVars(); //Retrieve the variable (cell) names from the tree
            var keys = variables.Keys;
            string thing = "";
            for (int i = 0; i < keys.Count; i++) //Look at each key
            {
                thing = keys.ElementAt(i);
                if (!mDependencies.ContainsKey(thing)) //If the dependencies dictionary does not contain thing
                {
                    HashSet<string> tempHash = new HashSet<string>(); //Push a new HashSet
                    tempHash.Add(cellName); //Push cellName to new hash set
                    mDependencies.Add(thing, tempHash); //Push thing and hash set to the dependency dictionary
                }
                else //If dependencies dictionary contains thing, add cellName to the dependency hash
                {
                    mDependencies[thing].Add(cellName);
                }
            }
        }

        public bool checkSelfReference(string cellName)
        {
            if (!mDependencies.ContainsKey(cellName)) //If this cell has no other cell referencing it
                return false;
            if (mDependencies[cellName].Contains(cellName)) //If this cell has references, check if it references itself
                return true;
            return false;
        }

        public bool checkCircularReference(string cellName)
        {
            string nextCell;
            if (!mDependencies.ContainsKey(cellName)) //If this cell has no other cell referencing it, it can't have a circular reference
                return false;
            for (int i = 0; i < mDependencies[cellName].Count; i++) //Look at all the dependencies of the current cell
            {
                nextCell = mDependencies[cellName].ElementAt(i); //Check the dependencies' dependencies
                if (!mDependencies.ContainsKey(nextCell))
                    break;
                if (mDependencies[nextCell].Contains(cellName)) //If a dependency refers back to the current cell, there is a circular reference
                    return true;
            }
            return false;
        }

        public void UpdateBGColor(Cell toColor)
        {
            CellPropertyChanged(toColor, new PropertyChangedEventArgs("BGColor")); //Update the UI
        }

        public void AddUndo(ICmd undoAction)
        {
            mUndoStack.Push(undoAction);
        }

        public void ExecuteUndo()
        {
            var pop = mUndoStack.Pop(); //Get the top element of the stack
            Cell currentCell = pop.GetCell();
            ICmd current = new MultiCmd();
            MultiCmd temp = current as MultiCmd;
            ICmd currentBGColor = new RestoreBGColor(currentCell, currentCell.BGColor);
            ICmd currentText = new RestoreText(currentCell, currentCell.Text);
            temp.Add(currentText);
            temp.Add(currentBGColor);
            AddRedo(current); //Push the current cell to the redo stack
            ICmd command = pop.Exec(); //Get the inverse action and apply to the cell
        }

        public bool undoIsEmpty()
        {
            if (mUndoStack.Count == 0)
                return true;
            return false;
        }

        public void AddRedo(ICmd redoAction)
        {
            mRedoStack.Push(redoAction);
        }

        public void ExecuteRedo()
        {
            var pop = mRedoStack.Pop();
            Cell currentCell = pop.GetCell();
            ICmd current = new MultiCmd();
            MultiCmd temp = current as MultiCmd;
            ICmd currentBGColor = new RestoreBGColor(currentCell, currentCell.BGColor);
            ICmd currentText = new RestoreText(currentCell, currentCell.Text);
            temp.Add(currentText);
            temp.Add(currentBGColor);
            AddUndo(current);
            ICmd command = pop.Exec();
        }

        public bool redoIsEmpty()
        {
            if (mRedoStack.Count == 0)
                return true;
            return false;
        }

        public void clearUndo()
        {
            mUndoStack.Clear();
        }

        public void clearRedo()
        {
            mRedoStack.Clear();
        }

        public void WriteToRandomCell()
        {
            Random random = new Random();
            for (int i = 0; i < 5; i++)
            {
                int randomRow = random.Next(0, 49); //Get a random row index
                int randomCol = random.Next(0, 25); //Get a random column index
                this.mSheet[randomRow, randomCol].Text = "Hello World!"; //Print Hello World to a random cell
            }
        }

        public enum alphaBool
        {
            NOT, UPPER, LOWER
        };

        private string toAlpha(string colIndex)
        {
            switch(colIndex)
            {
                case "0":
                    return "A";
                case "1":
                    return "B";
                case "2":
                    return "C";
                case "3":
                    return "D";
                case "4":
                    return "E";
                case "5":
                    return "F";
                case "6":
                    return "G";
                case "7":
                    return "H";
                case "8":
                    return "I";
                case "9":
                    return "J";
                case "10":
                    return "K";
                case "11":
                    return "L";
                case "12":
                    return "M";
                case "13":
                    return "N";
                case "14":
                    return "O";
                case "15":
                    return "P";
                case "16":
                    return "Q";
                case "17":
                    return "R";
                case "18":
                    return "S";
                case "19":
                    return "T";
                case "20":
                    return "U";
                case "21":
                    return "V";
                case "22":
                    return "W";
                case "23":
                    return "X";
                case "24":
                    return "Y";
                case "25":
                    return "Z";
            }
            return "";
        }
        private string toAlpha(int colIndex)
        {
            switch (colIndex)
            {
                case 0:
                    return "A";
                case 1:
                    return "B";
                case 2:
                    return "C";
                case 3:
                    return "D";
                case 4:
                    return "E";
                case 5:
                    return "F";
                case 6:
                    return "G";
                case 7:
                    return "H";
                case 8:
                    return "I";
                case 9:
                    return "J";
                case 10:
                    return "K";
                case 11:
                    return "L";
                case 12:
                    return "M";
                case 13:
                    return "N";
                case 14:
                    return "O";
                case 15:
                    return "P";
                case 16:
                    return "Q";
                case 17:
                    return "R";
                case 18:
                    return "S";
                case 19:
                    return "T";
                case 20:
                    return "U";
                case 21:
                    return "V";
                case 22:
                    return "W";
                case 23:
                    return "X";
                case 24:
                    return "Y";
                case 25:
                    return "Z";
            }
            return "";
        }
        public alphaBool isAlpha(char input)
        {
            switch (input)
            {
                case 'a':
                case 'b':
                case 'c':
                case 'd':
                case 'e':
                case 'f':
                case 'g':
                case 'h':
                case 'i':
                case 'j':
                case 'k':
                case 'l':
                case 'm':
                case 'n':
                case 'o':
                case 'p':
                case 'q':
                case 'r':
                case 's':
                case 't':
                case 'u':
                case 'v':
                case 'w':
                case 'x':
                case 'y':
                case 'z':
                    return alphaBool.LOWER;
                case 'A':
                case 'B':
                case 'C':
                case 'D':
                case 'E':
                case 'F':
                case 'G':
                case 'H':
                case 'I':
                case 'J':
                case 'K':
                case 'L':
                case 'M':
                case 'N':
                case 'O':
                case 'P':
                case 'Q':
                case 'R':
                case 'S':
                case 'T':
                case 'U':
                case 'V':
                case 'W':
                case 'X':
                case 'Y':
                case 'Z':
                    return alphaBool.UPPER;
                default:
                    return alphaBool.NOT;
            }
        }

        public void CopyBtoA()
        {
            int cols = 0;
            StringBuilder location = new StringBuilder();
            string[] columns = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            for (int rows = 0; rows < this.mRowCount && cols < this.mColumnCount; rows++)
            {
                location.Append("This is ");
                location.Append(columns[1]); //Appends the appropriate alphabetical column name
                location.Append((rows + 1).ToString()); //Appends the row number as a string
                this.mSheet[rows, cols+1].Text = location.ToString(); //Set the B#.Text to the string
                this.mSheet[rows, cols].Text = "=B" + rows+1; //Copy the B# cell to the A# cell
                Evaluate(this.mSheet[rows, cols]); //Evaluate the B# cell
                location.Clear(); //Clear the string so it can be used again
            }
        }

        public void Save(Stream toSave) //Saves the current spreadsheet to a stream
        {
            StreamWriter newFile = new StreamWriter(toSave);
            XDocument newXDoc = new XDocument();
            int i = 0, j = 0;
            string tempName = "";
            List<XElement> list = new List<XElement>();
            XElement root = new XElement("Spreadsheet"); //Begin the spreadsheet
            XElement cell, color, name, text, value, rowIndex, colIndex;
            for (; i < this.mRowCount; i++)
            {
                for (j = 0; j < this.mColumnCount; j++)
                {
                    if (this.mSheet[i, j].BGColor != 4294967295 || this.mSheet[i, j].Text != "")
                    {
                        tempName = toAlpha(j); //Get the column name as a string
                        tempName += this.mSheet[i, j].RowIndex.ToString(); //Get the row number as a string
                        name = new XElement("Name", tempName);
                        color = new XElement("BGColor", this.mSheet[i, j].BGColor);
                        text = new XElement("Text", this.mSheet[i, j].Text);
                        value = new XElement("Value", this.mSheet[i, j].Value);
                        rowIndex = new XElement("RowIndex", i);
                        colIndex = new XElement("ColIndex", j);
                        cell = new XElement("Cell", name, color, text, value, rowIndex, colIndex);
                        list.Add(cell);
                    }
                }
            }
            root.Add(list); //Add all the cells to the root
            newXDoc.Add(root); //Add the root to the XDoc
            newXDoc.Save(toSave); //Save the XDoc to the Stream
        }

        public void Load(Stream toLoad) //Loads a given stream
        {
            StreamReader newFile = new StreamReader(toLoad); //Load the stream into a stream reader
            string correctFormat = @"<?xml version=""1.0"" encoding=""utf-8""?>"; //All files saved with this format have this at the top
            string line = newFile.ReadLine(); //Get the first line of the text (XML version) and make sure its the correct kind
            if (line == correctFormat)
            {
                int row = -1, col = -1;
                uint color = 0;
                string cellName = "", text = "", value = "";
                while (!newFile.EndOfStream)
                {
                    line = newFile.ReadLine(); //Read the line
                    string[] name = line.Split('<', '>'); //Get the tag
                    string tag = name[1]; //The tag will always be at index[1] (i.e. name[0] = "", name[1] = "tag name", name[2] = "")
                    if (name.Length == 3) //There is only the tag name (i.e. <Spreadsheet>, </Spreadsheet>, <Cell>, </Cell>)
                    {
                        if (name[1] == "/Cell") //If the tag is the end of a cell, the info has been collected
                        {
                            this.mSheet[row, col].BGColor = color;
                            this.mSheet[row, col].Text = text;
                            this.mSheet[row, col].mTree = new ExpTree(text); //Evaluate Text to get the value
                        }
                    }
                    else //It is a cell's attribute (BGColor, Text, Name, Value, Row/Col index)
                    {
                        switch (name[1])
                        {
                            case "BGColor":
                                uint.TryParse(name[2], out color);
                                break;
                            case "Text":
                                text = name[2];
                                break;
                            case "Name":
                                cellName = name[2];
                                break;
                            case "Value":
                                value = name[2];
                                break;
                            case "RowIndex":
                                int.TryParse(name[2], out row);
                                break;
                            case "ColIndex":
                                int.TryParse(name[2], out col);
                                break;
                        }
                    }
                }
            }
        }
        public void Clear()
        {
            int i = 0, j = 0;
            for (; i < this.mRowCount; i++)
            {
                for (j = 0; j < this.mColumnCount; j++)
                {
                    this.mSheet[i, j].BGColor = 4294967295; //Set BGColor to white
                    this.mSheet[i, j].Text = ""; //Set text to empty string
                    this.mSheet[i, j].mTree = new ExpTree(""); //Set value to empty string
                }
            }
        }

    }


    public interface ICmd //Interfaces are not public by default, must use public access keyword
    {
        ICmd Exec();
        Cell GetCell();
    }

    public class RestoreText : ICmd
    {
        private Cell mCell;
        private string mText;
        public RestoreText(Cell c, string t)
        {
            mCell = c;
            mText = t;
        }

        public ICmd Exec()
        {
            ICmd inverse = new RestoreText(mCell, mCell.Text);
            mCell.Text = mText;
            return inverse;
        }

        public Cell GetCell()
        {
            return mCell;
        }

        public string GetText()
        {
            return mText;
        }
    }

    public class RestoreBGColor : ICmd
    {
        private Cell mCell;
        private uint mBGColor;

        public RestoreBGColor(Cell c, uint bgc)
        {
            mCell = c;
            mBGColor = bgc;
        }

        public ICmd Exec()
        {
            ICmd inverse = new RestoreBGColor(mCell, mCell.BGColor);
            mCell.BGColor = mBGColor;
            return inverse;
        }

        public Cell GetCell()
        {
            return mCell;
        }

        public uint GetBGColor()
        {
            return mBGColor;
        }
    }

    public class MultiCmd : ICmd
    {
        List<ICmd> mCmds;
        Cell c;
        string t;
        uint bgc;
        public MultiCmd()
        {
            mCmds = new List<ICmd>();
        }
   
        public ICmd Exec()
        {
            for (int i = 0; i < mCmds.Count; i++)
            {
                mCmds[i].Exec();
            }
            return this;
        }

        public Cell GetCell()
        {
            return c;
        }

        public void Add(ICmd newCmd)
        {
            mCmds.Add(newCmd); //Add to the mCmds list
            c = newCmd.GetCell(); //Update c
            t = c.Text; //Update t
            bgc = c.BGColor; //Update bgc
        }
    }

    public class ExpTree
    {
        public class Node //Base class
        {
            protected double mValue;
            public Node(double initValue = 0) { mValue = initValue; }
            public Node() { } //Empty constructor for OperatorNode to use
            public double Value
            {
                set { mValue = value; }
                get { return mValue; }
            }
            public virtual double Eval()
            {
                return mValue;
            }
        }

        public class ConstNode : Node //Constant number node
        {
            public ConstNode(double initValue) : base(initValue) { }
        }

        public class OperatorNode : Node
        {
            protected char mOperator;
            protected Node mLeft;
            protected Node mRight;
            public OperatorNode(char initOperator, Node initLeft, Node initRight)
            {
                mOperator = initOperator;
                mLeft = initLeft;
                mRight = initRight;
            }
            public Node Left
            {
                get { return mLeft; }
            }
            public Node Right
            {
                get { return mRight; }
            }
            public override double Eval()
            {
                switch (mOperator)
                {
                    case '+':
                        return mLeft.Eval() + mRight.Eval();
                    case '-':
                        return mLeft.Eval() - mRight.Eval();
                    case '*':
                        return mLeft.Eval() * mRight.Eval();
                    case '/':
                        return mLeft.Eval() / mRight.Eval();
                }
                return 0;
            }
        }

        public class VarNode : Node //Variable nodes
        {
            Dictionary<string, double> mCopy;
            protected string mName;
            public VarNode(string initName, Dictionary<string, double> treeDict)
            {
                mCopy = treeDict; //Get a copy of the ExpTree dictionary
                mName = initName;
                mValue = 0;
                if (!treeDict.ContainsKey(initName))
                    treeDict.Add(mName, mValue);
            }
            public string Name
            {
                set { mName = value; }
                get { return mName; }
            }
            
            public double Lookup (Dictionary<string, double> treeDict)
            {
                if (treeDict.ContainsKey(mName))
                    return treeDict[mName];
                else
                    return double.NaN;
            }

            public override double Eval()
            {
                return Lookup(mCopy);
            }
        }

        private Node mRoot;
        public Dictionary<string, double> mDict;

        public ExpTree(string expression)
        {
            this.mDict = new Dictionary<string, double>();
            this.mRoot = Compile(expression, mDict);
        }

        private static Node Compile (string expression, Dictionary<string, double> treeDict) //Called in the constructor to recursively build the expression tree
        {
            expression = expression.Replace(" ", ""); //Remove whitespace
            int lowIndex = GetLowOpIndex(expression);
            for (int i = expression.Length - 1; i >= 0; i--) //Start from the right side of the equation and look for operators
            {
                if (expression[0] == '(') //Check if the whole expression is inside parenthesis
                {
                    int numParenthesis = 1;
                    for (int j = 1; j < expression.Length; j++)
                    {
                        if (expression[j] == ')')
                            numParenthesis--;
                        if (expression[j] == '(')
                            numParenthesis++;
                        if (numParenthesis == 0 && j == expression.Length - 1 && expression[j] == ')')
                        {
                            return Compile(expression.Substring(1, expression.Length - 2), treeDict);
                        }
                    }
                }
                if (lowIndex == -1) //The lowest operator index is not an operator, it's a constant or a variable
                {
                    double num;
                    if (double.TryParse(expression, out num))
                        return new ConstNode(num);
                    else
                        return new VarNode(expression, treeDict);
                }
                if (lowIndex != -1) //The lowest operator index is an operator
                {
                    return new OperatorNode(expression[lowIndex], //Return a new operator node containing the lowest priority operator
                        Compile(expression.Substring(0, lowIndex), treeDict),
                        Compile(expression.Substring(lowIndex + 1), treeDict));
                }
                else
                    return null;
            }
            return BuildSimple(expression, treeDict); //Once the string is parsed, build the tree
        }

        public double Eval()
        {
            if (mRoot != null) { return this.mRoot.Eval(); }
            else
                return double.NaN;
        }

        public void SetVar(string varName, double varValue)
        {
            if (!mDict.ContainsKey(varName))
                mDict.Add(varName, varValue);
            else
                mDict[varName] = varValue;
        }

        public Dictionary<string, double> GetVars()
        {
            return mDict;
        }

        public static Node BuildSimple(string term, Dictionary<string, double> treeDict) //Builds a node, and depending on whether it can be parsed as a double or not
        {
            double num;
            if (double.TryParse(term, out num)) //See if the term is just a number
                return new ConstNode(num);
            return new VarNode(term, treeDict); //If the term can't be parsed as a double, then it's a variable
        }

        private static int GetLowOpIndex(string expression) //Gets the lowest priority operator
        {
            int parenthesis = 0, index = -1;
            for (int i = expression.Length - 1; i >= 0; i--) //Read expression right to left
            {
                switch (expression[i])
                {
                    case ')':
                        parenthesis--;
                        break;
                    case '(':
                        parenthesis++;
                        break;
                    case '+':
                    case '-':
                        if (parenthesis == 0) //If there are no parenthesis
                            return i;
                        break;
                    case '*':
                    case '/':
                        if (parenthesis == 0 && index == -1)
                            index = i; //Hold the rightmost index value in case this is the lowest priority operator
                        break;
                }
            }
            return index; //If there are no lower ( + or - ) operators
        }
    }
}