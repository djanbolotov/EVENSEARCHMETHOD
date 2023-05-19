using System;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using TextBox = System.Windows.Forms.TextBox;
using MessageBox = System.Windows.Forms.MessageBox;
using aziretParser;
using System.Diagnostics;
using System.Drawing;

namespace POCKETSEARCHMETHOD
{
    public partial class Form1 : Form
    {
        private const string nameOfExcel = @"\Zhanbolot_uulu_Askabek_LookingForOnePoint.xlsm";
        string inputFuncFX = "";
        decimal x0 = 0;
        decimal x1 = 0;
        decimal f0;
        decimal f1;
        decimal e_tol = 0;
        decimal h = 0;
        int k_max = 0;
        decimal t_max = 0;
        bool b = false;
        Application xls;
        Workbook book = null;
        Worksheet sheet = null;
        public Form1()
        {
            InitializeComponent();
            xls = new Application();
        }

        public int getSign(decimal number)
        {
            if (number < 0)
            {
                return -1;
            }
            else
            {
                return 1;
            }
        }

        public void OpenExcel()
        {
            if (!checkFunction(1)) return;
            string function;
            decimal startPoint;
            try
            {
                if (book == null)
                {
                    book = xls.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + nameOfExcel);
                }
                if (sheet == null)
                {
                    sheet = book.Sheets["Russian"];
                    sheet.Activate();
                }
                xls.Visible = true;
                function = Function.Text;
                if (InitialApproximation.Text != "" && InitialApproximation.Text != "-" && InitialApproximation.Text != "+" && InitialApproximation.Text != ".")
                {
                    startPoint = Decimal.Parse(InitialApproximation.Text);
                }
                else
                {
                    startPoint = 1;
                }

                sheet.Cells[4, 9] = startPoint;
                sheet.Cells[2, 1] = "f(x)=" + Function.Text;
            }
            catch
            {
                book = xls.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + nameOfExcel);
                sheet = book.Sheets["Russian"];
                sheet.Activate();
                xls.Visible = true;
                function = Function.Text;
                if (InitialApproximation.Text != "" && InitialApproximation.Text != "-" && InitialApproximation.Text != "+" && InitialApproximation.Text != ".")
                {
                    startPoint = Decimal.Parse(InitialApproximation.Text);
                }
                else
                {
                    startPoint = 1;
                }

                sheet.Cells[4, 9] = startPoint;
                sheet.Cells[2, 1] = "f(x)=" + Function.Text;
            }

            StringBuilder builder = new StringBuilder(function);
            builder.Replace("exp", ":");
            builder.Replace("x", "D4");
            builder.Replace(":", "exp");
            function = builder.ToString();
            sheet.Range["E4:E10003"].Value = "=" + function;
        }

        private bool parseTry(TextBox t, String type)
        {
            try
            {
                if (type == "Decimal")
                    Decimal.Parse(t.Text, System.Globalization.NumberStyles.Float);
                else if (type == "Integer")
                    int.Parse(t.Text);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void Clean(Control control)
        {
            foreach (var  element in control.Controls)
            {
                switch (element.GetType().Name)
                {
                    case "TextBox":
                        ((TextBox)element).Text = String.Empty;
                        break;
                    case "RadioButton":
                        ((RadioButton)element).Checked = false;
                        break;
                    case "RichTextBox":
                        ((RichTextBox)element).Text = String.Empty;
                        break;
                    case "ProgressBar":
                        ((System.Windows.Forms.ProgressBar)element).Value = 0;
                        break;
                    case "GroupBox":
                        Clean((Control)element);
                        break;
                    default:
                        break;
                }
            }
        }


        private bool IsOKForDecimalTextBox(char theCharacter, TextBox theTextBox, bool positive)
        {
            if (!char.IsControl(theCharacter) && !char.IsDigit(theCharacter) && (theCharacter != ',')
                && (theCharacter != '-') && (theCharacter != '+') && (theCharacter != 'E') && (theCharacter != 'e'))
            {
                return false;
            }
            if(positive && theCharacter == '-' && (theTextBox.Text.IndexOf('E') == -1 && theTextBox.Text.IndexOf('e') == -1))
            {
                return false;
            }
            if (theCharacter == ',' && theTextBox.Text.IndexOf(',') > -1)
            {
                return false;
            }
            if (theCharacter == 'e' && (theTextBox.Text.IndexOf('e') > -1 || theTextBox.Text.IndexOf('E') > -1))
            {
                return false;
            }
            if (theCharacter == 'E' && (theTextBox.Text.IndexOf('E') > -1 || theTextBox.Text.IndexOf('e') > -1))
            {
                return false;
            }
            if (theCharacter == '-' && (theTextBox.Text.IndexOf('-') > -1 || theTextBox.Text.IndexOf('+') > -1))
            {
                return false;
            }
            if (theCharacter == '+' && (theTextBox.Text.IndexOf('+') > -1 || theTextBox.Text.IndexOf('-') > -1))
            {
                return false;
            }
            if (((theCharacter == '-') || (theCharacter == '+')) && (theTextBox.SelectionStart != 0 && (theTextBox.Text.IndexOf('E') == -1 && theTextBox.Text.IndexOf('e') == -1)))
            {
                return false;
            }
            if ((char.IsDigit(theCharacter) || (theCharacter == ',')) && ((theTextBox.Text.IndexOf('-') > -1) 
                || (theTextBox.Text.IndexOf('+') > -1)) && theTextBox.SelectionStart == 0)
            {
                return false;
            }
            return true;
        }

        public decimal F(decimal x)
        {
            decimal result;
            result = aziretParser.Computer.Compute(inputFuncFX, x);
            return result;
        }
        private void button4_Click(object sender, EventArgs e)
        {
            OpenExcel();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Clean(this);
        }

        private void InitialApproximation_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '.')
            {
                e.KeyChar = ',';
            }
            e.Handled = !IsOKForDecimalTextBox(e.KeyChar, InitialApproximation, false);
        }

        private void Tolerance_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '.')
            {
                e.KeyChar = ',';
            }
            e.Handled = !IsOKForDecimalTextBox(e.KeyChar, Tolerance, true);
        }

        private void SearchStep_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '.')
            {
                e.KeyChar = ',';
            }
            e.Handled = !IsOKForDecimalTextBox(e.KeyChar, SearchStep, true);
        }

        private String checkParse()
        {
            String errorMessage = "";

            if (!parseTry(InitialApproximation, "Decimal"))
            {
                errorMessage += "Invalid value of the field x0 (the starting point of the approximation)! Change the input and perform the calculation!\n\n";
            }
            else
            {
                x0 = Decimal.Parse(InitialApproximation.Text, System.Globalization.NumberStyles.Float);
            }

            if (!parseTry(SearchStep, "Decimal"))
            {
                errorMessage += "Invalid value of the field search step! Change the input and perform the calculation!\n\n";
            }
            else
            {
                h = Decimal.Parse(SearchStep.Text, System.Globalization.NumberStyles.Float);
            }

            if (parseTry(Tolerance, "Decimal"))
            {
                e_tol = Decimal.Parse(Tolerance.Text, System.Globalization.NumberStyles.Float);
            }
            else
            {
                errorMessage += "Invalid value of the Tolerance(e) field (entered tolerance)! Change the input and perform the calculation!\n\n";
            }


            if (!parseTry(LimitOfIterations, "Integer"))
            {
                errorMessage += "Invalid value of the field limit of iterations! Change the input and perform the calculation!\n\n";
            }
            else
            {
                k_max = Int32.Parse(LimitOfIterations.Text);
            }

            if (!parseTry(LimitOfTime, "Decimal"))
            {
                errorMessage += "Invalid value of the field limit of time! Change the input and perform the calculation!\n\n";
            }
            else
            {
                t_max = Decimal.Parse(LimitOfTime.Text, System.Globalization.NumberStyles.Float);
            }


            return errorMessage;
        }

        public bool fullCheck()
        {
            bool check = false;
            if (Function.Text == "" || InitialApproximation.Text == "" ||
                Tolerance.Text == "" || LimitOfIterations.Text == "" ||
                LimitOfTime.Text == "" || SearchStep.Text == "")
            {
                MessageBox.Show("All fields must be filled in! Enter the missing information and make the calculation!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                if (checkParse() != "")
                {
                    MessageBox.Show(checkParse(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    if (isRigth() && checkFunction(x0))
                    {
                        check = true;
                    }
                }
            }
            return check;
        }

        private void LimitOfTime_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '.')
            {
                e.KeyChar = ',';
            }
            e.Handled = !IsOKForDecimalTextBox(e.KeyChar, LimitOfTime, true);
        }

        private void LimitOfIterations_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((int)e.KeyChar == (int)48 && LimitOfIterations.Text == "")
            {
                e.Handled = true;
                return;
            }
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }
        public string getComparisonSign(decimal a, decimal b)
        {
            if (a > b)
            {
                return ">";
            }
            else if (a < b)
            {
                return "<";
            }
            else
            {
                return "=";
            }
        }

        private bool isRigth()
        {
            bool valid = true;
            if (e_tol <= 0)
            {
                MessageBox.Show("The value of the tolerance field must be greater than 0! Change the input and perform the calculation!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                valid = false;
            }
            if(h <= 0)
            {
                MessageBox.Show("The value of the search step field must be greater than 0! Change the input and perform the calculation!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                valid = false;
            }
            if(h > e_tol)
            {
                MessageBox.Show("The value of the search step field must be equal or less than tolerance field! Change input and perform the calculation!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                valid = false;
            }
            if (k_max <= 0)
            {
                MessageBox.Show("The value of the limit of iterations field must be greater than 0! Change the input and perform the calculation!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                valid = false;
            }
            if (t_max <= 0)
            {
                MessageBox.Show("The value of the limit of time field must be greater than 0! Change the input and perform the calculation!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                valid = false;
            }
            if (!(Maximum.Checked || Minimum.Checked))
            {
                MessageBox.Show("Please select search option maximum or minimum.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                valid = false;
            }
            if (valid)
            {
                return true;
            }
            return false;
        }

        private bool checkFunction(decimal x0)
        {
            inputFuncFX = Function.Text;

            if (inputFuncFX == "" || inputFuncFX.IndexOf('x') == -1)
            {
                MessageBox.Show("The function is entered incorrectly! Change the input and perform the calculation!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Clean(this);
                return false;
            }
            try
            {
                if (inputFuncFX.Contains("log") && x0 <= 0 || inputFuncFX.Contains("ln") && x0 <= 0)
                {
                    MessageBox.Show("If you entered function with 'log' or 'ln' value of X0 must greater than zero!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                else
                {
                    decimal F1 = F(x0);
                    return true;
                }
            }
            catch
            {
                MessageBox.Show("The function or initial approximation is entered incorrectly! Change the input and perform the calculation!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Clean(this);
                return false;
            }
        }

        public bool MaxOrMin(decimal f0, decimal f1)
        {
            if (Maximum.Checked)
            {
                return f0 >= f1;
            }
            return f0 <= f1;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("1) Choose a function or write your's on field 'Function'\n" +
                     "2) Click on the button 'Show function graph'\n" +
                     "3) In the opened file select the values for a or b,\n" +
                     "then save the document and return to the program\n" +
                     "4) If you need 'a' value to insert,\n" +
                     "click the button 'Set 'a' like 'X0'' or write your's\n" +
                     "5) Enter tolerance\n" +
                     "6) Enter search step\n" +
                     "8) Enter limit of time in sec\n" +
                     "9) Enter limit of iterations \n" +
                     "10) Select search parameter\n" +
                     "Then click the button 'Run Method'.", "Information",
                     MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (book == null)
                {
                    book = xls.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + nameOfExcel);
                }
                if (sheet == null)
                {
                    sheet = book.Sheets["Russian"];
                    sheet.Activate();
                }
                book.Save();
                InitialApproximation.Text = sheet.Cells[4, 9].Value.ToString();
            }
            catch
            {
                book = xls.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + nameOfExcel);
                sheet = book.Sheets["Russian"];
                sheet.Activate();
                book.Save();
                InitialApproximation.Text = sheet.Cells[4, 9].Value.ToString();
            }
            xls.Visible = false;
            book = null;
            sheet = null;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            inputFuncFX = "";
            x0 = 0;
            x1 = 0;
            f0 = 0;
            f1 = 0;
            e_tol = 0;
            k_max = 0;
            t_max = 0;

            string extremium;

            if (fullCheck())
            {
                xls.Visible = false;
                book = null;
                sheet = null;
                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();
                Clean(groupBox2);
                validation.Text = String.Empty;
                progressBar1.Value = 0;

                decimal fplusTol;
                decimal fminusTol;

                f0 = F(x0);

                x1 = x0 + e_tol;

                f1 = F(x1);

                int k = 0;

                if (Maximum.Checked)
                {
                    extremium = "maximizer";
                }
                else
                {
                    extremium = "minimizer";
                }

                progressBar1.Value = 0;

                while (true)
                {
                    k = k + 1;

                    progressBar1.Visible = true;
                    progressBar1.Maximum = (int)(k + 0.00000001);
                    progressBar1.Value = k;

                    if (k > k_max)
                    {
                        stopwatch.Stop();
                        f1 = F(x1);
                        fminusTol = F(x1 - e_tol);
                        fplusTol = F(x1 + e_tol);
                        DialogResult result = MessageBox.Show("Iteration limit reached. Do you want to add iterations?",
                            "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                        if (result == DialogResult.Yes)
                        {
                            k_max += k_max;
                            LimitOfIterations.Text = k_max.ToString();
                        }
                        else
                        {
                            k--;
                            validation.Text += "Result X* not found because of limit of iterations = " + k_max + "." +
                                "\nSince the following condition is false, namely:" +
                                "\nSign(f(X*)-f(X*+Tolerance)) = " + getSign(f1 - fplusTol) + " and Sign(f(X*)-f(X*-Tolerance)) = " + getSign(f1 - fminusTol) + "!" +
                                "\nResult X* is not " + extremium + " of the function.";
                            validation.ForeColor = Color.Red;

                            FillResult(x1.ToString("F28"), k.ToString(), getError(Tolerance, Math.Abs(x1 - x0)), fminusTol.ToString("F28"), fplusTol.ToString("F28"), f1.ToString("F28"), (f1 - fplusTol).ToString("F28"), (f1 - fminusTol).ToString("F28"));
                            absError.Text = getError(Tolerance, Math.Abs(x1 - x0));

                            DialogResult answer = MessageBox.Show("Result X* not found because of maximum limit of iterations = " + k_max + "." +
                            "\nSince the following condition is false, namely:" +
                            "\nSign(f(X*)-f(X*+Tolerance)) = " + getSign(f1 - fplusTol) + " and Sign(f(X*)-f(X*-Tolerance)) = " + getSign(f1 - fminusTol) + "!" +
                            "\nResult X* is not " + extremium + " of the function." +
                            "\n\nYou probably entered the values of 'a' and 'b' range incorrectly on Ecxel!" +
                            "\nSince the program is looking for an extremum only in the range 'a' and 'b'." +
                            "\nYou need to open the graph and select the correct points [a;b]!" +
                            "\n\nDo you want to open file?", "Error", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                            if (answer == DialogResult.Yes)
                            {
                                OpenExcel();
                            }
                            break;
                        }
                        stopwatch.Start();
                    }

                    if (stopwatch.ElapsedMilliseconds >= t_max * 1000)
                    {
                        stopwatch.Stop();
                        f1 = F(x1);
                        fminusTol = F(x1 - e_tol);
                        fplusTol = F(x1 + e_tol);
                        DialogResult result = MessageBox.Show("Time limit reached. Do you want to add time?",
                            "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                        if (result == DialogResult.Yes)
                        {
                            t_max += t_max;
                            LimitOfTime.Text = t_max.ToString();
                        }
                        else
                        {
                            validation.Text += "Result X* not found because of limit of time = " + t_max + " sec." +
                                "\nSince the following condition is false, namely:" +
                                "\nSign(f(X*)-f(X*+Tolerance)) = " + getSign(f1 - fplusTol) + " and Sign(f(X*)-f(X*-Tolerance)) = " + getSign(f1 - fminusTol) + "!" +
                                "\nResult X* is not " + extremium + " of the function.";
                            validation.ForeColor = Color.Red;

                            FillResult(x1.ToString("F28"), k.ToString(), getError(Tolerance, Math.Abs(x1 - x0)), fminusTol.ToString("F28"), fplusTol.ToString("F28"), f1.ToString("F28"), (f1 - fplusTol).ToString("F28"), (f1 - fminusTol).ToString("F28"));
                            absError.Text = getError(Tolerance, Math.Abs(x1 - x0));

                            DialogResult answer = MessageBox.Show("Result X* not found because of maximum time limit = " + t_max + " sec." +
                            "\nSince the following condition is false, namely:" +
                            "\nSign(f(X*)-f(X*+Tolerance)) = " + getSign(f1 - fplusTol) + " and Sign(f(X*)-f(X*-Tolerance)) = " + getSign(f1 - fminusTol) + "!" +
                            "\nResult X* is not " + extremium + " of the function." +
                            "\n\nYou probably entered the values of 'a' and 'b' range incorrectly on Ecxel!" +
                            "\nSince the program is looking for an extremum only in the range 'a' and 'b'." +
                            "\nYou need to open the graph and select the correct points [a;b]!" +
                            "\n\nDo you want to open file?", "Error", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                            if (answer == DialogResult.Yes)
                            {
                                OpenExcel();
                            }
                            break;
                        }
                        stopwatch.Start();
                    }

                    if (MaxOrMin(f0, f1))
                    {
                        x1 = x0;
                        f1 = f0;

                        Clean(this);

                        validation.Text += "The program did not run a single iteration!";
                        validation.ForeColor = Color.Red;

                        DialogResult result = MessageBox.Show("The program did not run a single iteration. The starting point is " + extremium + " of function or is to the right of it or search step (h) is too small. You need to open the graph and select the correct point! \n\nDo you want to open file?", "Error", MessageBoxButtons.YesNo, MessageBoxIcon.Error);

                        if (result == DialogResult.Yes)
                        {
                            OpenExcel();
                        }
                        return;
                    }
                    else
                    {
                        x0 = x1;
                        f0 = f1;
                        x1 = x0 + h;
                        f1 = F(x1);
                    }

                    fminusTol = F(x1 - e_tol);
                    fplusTol = F(x1 + e_tol);

                    if (Math.Abs(x1 - x0) != 0)
                    {
                        if(extremium == "minimizer")
                        {
                            if ((f1 <= fminusTol && f1 <= fplusTol))
                            {
                                FillResult(x1.ToString("F28"), k.ToString(), getError(Tolerance, Math.Abs(x1 - x0)), fminusTol.ToString("F28"), fplusTol.ToString("F28"), f1.ToString("F28"), (f1 - fplusTol).ToString("F28"), (f1 - fminusTol).ToString("F28"));

                                validation.Text += "Since the following condition is true, namely:" +
                                        "\nSign(f(X*)-f(X*+Tolerance)) = " + getSign(f1 - fplusTol) + " and Sign(f(X*)-f(X*-Tolerance)) = " + getSign(f1 - fminusTol) + "!" +
                                        "\nResult X* is " + extremium + " of the function. It has been found with the error = " + getError(Tolerance, Math.Abs(x1 - x0)) + ". This is less than or equal to given Tolerance!";

                                validation.ForeColor = Color.Green;

                                break;
                            }
                        }
                        else
                        {
                            if (f1 >= fminusTol && f1 >= fplusTol)
                            {
                                FillResult(x1.ToString("F28"), k.ToString(), getError(Tolerance, Math.Abs(x1 - x0)), fminusTol.ToString("F28"), fplusTol.ToString("F28"), f1.ToString("F28"), (f1 - fplusTol).ToString("F28"), (f1 - fminusTol).ToString("F28"));

                                validation.Text += "Since the following condition is true, namely:" +
                                        "\nSign(f(X*)-f(X*+Tolerance)) = " + getSign(f1 - fplusTol) + " and Sign(f(X*)-f(X*-Tolerance)) = " + getSign(f1 - fminusTol) + "!" +
                                        "\nResult X* is " + extremium + " of the function. It has been found with the error = " + getError(Tolerance, Math.Abs(x1 - x0)) + ". This is less than or equal to given Tolerance!";

                                validation.ForeColor = Color.Green;

                                break;
                            }
                        }
                    }
                }

                stopwatch.Stop();
                elapsedtime.Text = stopwatch.ElapsedMilliseconds / 1000.0 + " sec";

                timer1.Enabled = true;
                timer1.Start();
            }
        }

        public void FillResult(string solution, string iterations, string resultTolerance, string fminustol, string fplustol, string fxvalue, string fminusplus, string fminusminus)
        {
            ResultX.Text = solution;
            countofiterations.Text = iterations;
            fxplustolerance.Text = fplustol;
            fxminustolerance.Text = fminustol;
            fxminusplustolerance.Text = fminusplus;
            fxminusminustolerance.Text = fminusminus;
            fx.Text = fxvalue;
            absError.Text = resultTolerance;
        }

        public string getError(TextBox tol, decimal error)
        {
            Console.WriteLine(tol);
            if (tol.Text.Contains("E"))
            {
                return error.ToString("0E0");
            }
            else if (tol.Text.Contains("e"))
            {
                return error.ToString("0e0");
            }
            else
            {
                return error.ToString();
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            progressBar1.Value = 0;
            timer1.Enabled = false;
            timer1.Stop();
        }

        private void Function_TextChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void InitialApproximation_TextChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void Tolerance_TextChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void SearchStep_TextChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void ParametrR_TextChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void LimitOfTime_TextChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void LimitOfIterations_TextChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void Maximum_CheckedChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void Minimum_CheckedChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            xls.Quit();
        }
    }
}
