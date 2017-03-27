/*
 * Courtney Snyder
 * CptS 321, Homework 7
 * Last Updated: 3/10/2017
 * Description: This is where the user interface for the spreadsheet app is at. The spreadsheet is represented using a dataGridView
 * object.
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using SpreadsheetEngine; //Using the class library SpreadsheetEngine

namespace Spreadsheet_CSnyder
{
    public partial class Form1 : Form
    {
        public event DataGridViewCellCancelEventHandler CellBeginEdit = delegate { };
        public event DataGridViewCellEventHandler CellEndEdit = delegate { };

        public Form1()
        {
            InitializeComponent();
        }
        Spreadsheet test = new Spreadsheet(50, 26);
        
        private void Form1_Load(object sender, EventArgs e)
        {
            this.dataGridView.Columns.Clear();
            string[] columnNames = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"};
            for (int colIndex = 0; colIndex < columnNames.Count(); colIndex++)
                this.dataGridView.Columns.Add(columnNames.ElementAt(colIndex).ToString(), columnNames.ElementAt(colIndex).ToString()); //Add 1 column and give it a header value
            this.dataGridView.Rows.Add(50); //Add 50 rows
            for (int rowIndex = 0; rowIndex < 50; rowIndex++)
                this.dataGridView.Rows[rowIndex].HeaderCell.Value = (rowIndex + 1).ToString(); //Update each row's header
            test.CellPropertyChanged += OnCellPropertyChanged; //Subscribe to CellPropertyChanged
        }

        protected void OnCellPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            Cell currentCell = sender as Cell;

            if (e.PropertyName == "Value")
            {
                //ICmd action = new RestoreText(currentCell, currentCell.Text);
                //test.AddUndo(action);
                dataGridView[currentCell.mColumnIndex, currentCell.mRowIndex].Value = currentCell.Value; //Display the cell's value to the UI
            }
            if (e.PropertyName == "BGColor")
            {
                //ICmd action = new RestoreBGColor(currentCell, currentCell.BGColor);
                //test.AddUndo(action);
                Color newColor = Color.FromArgb((int)currentCell.BGColor);
                dataGridView[currentCell.mColumnIndex, currentCell.mRowIndex].Style.BackColor = newColor; //Update the back color of each dataGridView cell to the color of the Cell
            }
        }

        private void dataGridView_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            Cell currentCell = this.test.GetCell(e.RowIndex, e.ColumnIndex);
            dataGridView[e.ColumnIndex, e.RowIndex].Value = currentCell.Text;
            Color newColor = Color.FromArgb((int)currentCell.BGColor);
            dataGridView[e.ColumnIndex, e.RowIndex].Style.BackColor = newColor;
            ICmd text = new RestoreText(currentCell, currentCell.Text);
            ICmd bgcolor = new RestoreBGColor(currentCell, currentCell.BGColor);
            ICmd multi = new MultiCmd();
            MultiCmd temp = multi as MultiCmd;
            temp.Add(text);
            temp.Add(bgcolor);
            test.AddUndo(multi);
        }

        private void dataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            Cell currentCell = this.test.GetCell(e.RowIndex, e.ColumnIndex);
            if (dataGridView[e.ColumnIndex, e.RowIndex].Value != null)
                currentCell.Text = dataGridView[e.ColumnIndex, e.RowIndex].Value.ToString();
            Color newColor = dataGridView[e.ColumnIndex, e.RowIndex].Style.BackColor; //Get the color from dataGridView
            currentCell.BGColor = (uint)newColor.ToArgb();
        }

        private void changeColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                Color newColor = colorDialog.Color; //Save the user selected color
                dataGridView.BackgroundColor = newColor; //Save the new color in dataGridView.BackgroundColor
                var currentCells = dataGridView.SelectedCells; //Get the highlighted cells
                for (int i = 0; i < currentCells.Count; i++)
                {
                    Cell currentCell = test.GetCell(currentCells[i].RowIndex, currentCells[i].ColumnIndex);
                    currentCell.BGColor = (uint)newColor.ToArgb(); //Update the background color of each selected cell to the saved color the user selected
                    ICmd text = new RestoreText(currentCell, currentCell.Text);
                    ICmd bgcolor = new RestoreBGColor(currentCell, currentCell.BGColor);
                    ICmd multi = new MultiCmd();
                    MultiCmd temp = multi as MultiCmd;
                    temp.Add(text);
                    temp.Add(bgcolor);
                    test.AddUndo(multi);
                }
            }
        }

        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!test.undoIsEmpty())
            {
                test.ExecuteUndo();
            }
        }

        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!test.redoIsEmpty())
            {
                test.ExecuteRedo();
            }
        }

        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.openFileDialog.ShowDialog(); //Get the file the user wants to load
            if (openFileDialog.FileName != "")
            {
                test.Clear(); //Clear the current contents of the sheet
                test.clearUndo(); //Clear the undo stack
                test.clearRedo(); //Clear the redo stack
                Stream loadStream = this.openFileDialog.OpenFile(); //Open the file as a stream
                this.test.Load(loadStream); //Load the stream into the spreadsheet
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.saveFileDialog.ShowDialog(); //Get the filename and path the user wants to save to
            if (saveFileDialog.FileName != "")
            {
                Stream saveStream = this.saveFileDialog.OpenFile();
                this.test.Save(saveStream); //Save the spreadsheet
            }
        }

        /*
        private void button_Click(object sender, EventArgs e)
        {
            this.test.WriteToRandomCell(); //Writes "Hello World" to 50 random cells
            this.test.CopyBtoA(); //Writes "This is B#" to all the B cells and copies them to all the A cells
        } 
        */
    }
}