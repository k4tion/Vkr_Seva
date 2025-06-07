using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace VKRShevelev
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private List<Project> projects = new List<Project>();

        private string GetYearLabel(double years)
        {
            int y = (int)Math.Round(years);

            if (y % 100 >= 11 && y % 100 <= 14)
                return "лет";

            switch (y % 10)
            {
                case 1: return "год";
                case 2:
                case 3:
                case 4: return "года";
                default: return "лет";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var tempProject = CalculateProject("Временный проект");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                string projectName = textBox12.Text.Trim();
                if (string.IsNullOrEmpty(projectName))
                    throw new Exception("Введите название проекта.");

                if (projects.Any(p => p.Name.Equals(projectName, StringComparison.OrdinalIgnoreCase)))
                    throw new Exception("Проект с таким названием уже существует.");

                var project = CalculateProject(projectName);
                projects.Add(project);
                UpdateProjectGrid();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = projects.Select(p => new
            {
                Название = p.Name,
                ИтоговаяОценка = Math.Round(p.TotalScore, 2),
                ДисконтируемыйСрокОкупаемости = $"{Math.Round(p.PaybackPeriod, 2)} {GetYearLabel(p.PaybackPeriod)}"
            }).ToList();

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.ClearSelection();
        }

        private void UpdateProjectGrid()
        {
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = projects.Select(p => new
            {
                Название = p.Name,
                ИтоговаяОценка = Math.Round(p.TotalScore, 2),
                ДисконтируемыйСрокОкупаемости = $"{Math.Round(p.PaybackPeriod, 2)} {GetYearLabel(p.PaybackPeriod)}"
            }).ToList();

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.ClearSelection();
        }

        private Project CalculateProject(string projectName)
        {
            try
            {
                // Валидация целочисленных оценок
                if (!int.TryParse(textBox1.Text, out int marginScore))
                    throw new Exception("Поле 'Маржинальность продукции' должно содержать целое число.");
                if (!int.TryParse(textBox2.Text, out int paybackScore))
                    throw new Exception("Поле 'Дисконтируемый срок окупаемости' должно содержать целое число.");
                if (!int.TryParse(textBox3.Text, out int techFeasibilityScore))
                    throw new Exception("Поле 'Технологическая реализуемость' должно содержать целое число.");
                if (!int.TryParse(textBox4.Text, out int ecoDemandScore))
                    throw new Exception("Поле 'Запрос на экологичное производство' должно содержать целое число.");
                if (!int.TryParse(textBox5.Text, out int stateSupportScore))
                    throw new Exception("Поле 'Государственная принадлежность' должно содержать целое число.");
                if (!int.TryParse(textBox6.Text, out int marketDemandScore))
                    throw new Exception("Поле 'Объём платежеспособного спроса' должно содержать целое число.");

                // Валидация долей от 0 до 1
                if (!double.TryParse(textBox7.Text, out double techFeasibilityWeight) || techFeasibilityWeight < 0 || techFeasibilityWeight > 1)
                    throw new Exception("Поле 'Значение Технологическая реализуемость' должно быть числом от 0 до 1.");
                if (!double.TryParse(textBox8.Text, out double ecoDemandWeight) || ecoDemandWeight < 0 || ecoDemandWeight > 1)
                    throw new Exception("Поле 'Значение Запрос на экологичное производство' должно быть числом от 0 до 1.");
                if (!double.TryParse(textBox9.Text, out double stateSupportWeight) || stateSupportWeight < 0 || stateSupportWeight > 1)
                    throw new Exception("Поле 'Значение Государственная принадлежность' должно быть числом от 0 до 1.");
                if (!double.TryParse(textBox10.Text, out double marketDemandWeight) || marketDemandWeight < 0 || marketDemandWeight > 1)
                    throw new Exception("Поле 'Значение Объём платежеспособного спроса' должно быть числом от 0 до 1.");

                // CAPEX и ставка дисконтирования
                if (!double.TryParse(capex.Text, out double investment))
                    throw new Exception("Поле 'Инвестиции' должно содержать число.");
                if (!double.TryParse(rate.Text, out double discountRate) || discountRate < 0 || discountRate > 100)
                    throw new Exception("Ставка дисконтирования должна быть от 0 до 100%.");
                discountRate /= 100;

                // Цена, себестоимость, объём продаж
                if (!double.TryParse(price.Text, out double unitPrice))
                    throw new Exception("Поле 'Цена' должно содержать число.");
                if (!double.TryParse(cost.Text, out double unitCost))
                    throw new Exception("Поле 'Себестоимость' должно содержать число.");
                if (!double.TryParse(volume.Text, out double salesVolume))
                    throw new Exception("Поле 'Объём продаж' должно содержать число.");

                // Горизонт расчёта
                if (!int.TryParse(textBox11.Text, out int planningHorizon))
                    throw new Exception("Поле 'Горизонт расчёта' должно содержать целое число.");

                // Расчёт маржинальности
                double marginRate = (unitPrice - unitCost) / unitPrice;
                label33.Text = $"{Math.Round(marginRate, 2)}";

                // Расчёт дисконтированного срока окупаемости
                double annualProfit = (unitPrice - unitCost) * salesVolume;
                double discountedCashFlow = 0;
                double paybackPeriod = -1;

                for (int year = 1; year <= planningHorizon; year++)
                {
                    double discountedProfit = annualProfit / Math.Pow(1 + discountRate, year);
                    discountedCashFlow += discountedProfit;

                    if (discountedCashFlow >= investment)
                    {
                        double previousCashFlow = discountedCashFlow - discountedProfit;
                        double fractionOfYear = (investment - previousCashFlow) / discountedProfit;
                        paybackPeriod = (year - 1) + fractionOfYear;
                        break;
                    }
                }

                label34.Text = paybackPeriod > 0
                    ? $"{Math.Round(paybackPeriod, 2)} {GetYearLabel(paybackPeriod)}"
                    : "Не окупается";

                label35.Text = paybackPeriod > 0
                    ? $"{Math.Round(planningHorizon / paybackPeriod, 2)}"
                    : "∞";

                // Комплексная оценка
                double scoreMargin = 0.4 * marginScore * marginRate;
                double scorePayback = 0.15 * paybackScore * paybackPeriod;
                double scoreTech = 0.15 * techFeasibilityScore * techFeasibilityWeight;
                double scoreEco = 0.1 * ecoDemandScore * ecoDemandWeight;
                double scoreState = 0.1 * stateSupportScore * stateSupportWeight;
                double scoreMarket = 0.1 * marketDemandScore * marketDemandWeight;

                label13.Text = $"{Math.Round(scoreMargin, 2)}";
                label14.Text = $"{Math.Round(scorePayback, 2)}";
                label15.Text = $"{Math.Round(scoreTech, 2)}";
                label16.Text = $"{Math.Round(scoreEco, 2)}";
                label17.Text = $"{Math.Round(scoreState, 2)}";
                label18.Text = $"{Math.Round(scoreMarket, 2)}";

                double total = scoreMargin + scorePayback + scoreTech + scoreEco + scoreState + scoreMarket;
                label19.Text = $"{Math.Round(total, 2)}";

                label26.Text = total > 4 ? "Проект рекомендуется к реализации" : "Проект нецелесообразен к реализации";
                label26.ForeColor = total > 4 ? Color.Green : Color.Red;

                return new Project
                {
                    Name = projectName,
                    MarginScore = scoreMargin,
                    PaybackScore = scorePayback,
                    TechnologyScore = scoreTech,
                    DemandScore = scoreEco,
                    StrategyScore = scoreState,
                    VolumeScore = scoreMarket,
                    PaybackPeriod = paybackPeriod
                };
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

    }

    public class Project
    {
        public string Name { get; set; }
        public double MarginScore { get; set; }
        public double PaybackScore { get; set; }
        public double TechnologyScore { get; set; }
        public double DemandScore { get; set; }
        public double StrategyScore { get; set; }
        public double VolumeScore { get; set; }
        public double TotalScore => MarginScore + PaybackScore + TechnologyScore + DemandScore + StrategyScore + VolumeScore;
        public double PaybackPeriod { get; set; }
    }
}
