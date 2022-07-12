using System; //подключение необходимых библиотек
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace Программа_с_формами
{
    public partial class Form1 : Form
    {
        public static Microsoft.Office.Interop.Excel.Application excel; //создание компонентов для работы с Excel
        public static Microsoft.Office.Interop.Excel.Workbook wb; //wb - рабочая книга Excel
        public static Microsoft.Office.Interop.Excel.Worksheet ws; //ws - рабочая страница Excel

        public Form1() //конструктор для класса Form1
        {
            InitializeComponent(); //инициализация компонентов
            textBox1.Enabled = false; //необходимые компонетны закрываются для использования
            listBox1.Visible = false;
            button3.Enabled = false;
            button4.Enabled = false;
            label3.Visible = false;
            label4.Visible = false;
            label5.Visible = false;
        }

        public void button1_Click(object sender, EventArgs e) //действия при нажатии на кнопку "Загрузить таблицу"
        {
            if (excel != null)
            {
                excel.Quit(); // закрытие Excel файла
                excel = null; // обнуление всех ссылок для корректного закрытия Excel файла
                wb = null;
                ws = null;
            }

            button3.Enabled = false;
            button4.Enabled = false;
            comboBox1.Items.Clear(); //очистка поля со списком ComboBox1
            comboBox2.Items.Clear(); //очистка поля со списком ComboBox2

            openFileDialog1.Filter = "Excel файлы(*.xlsx)|*.xlsx"; //фильтр для открытия файлов
            openFileDialog1.ShowDialog(); //открывается диалог выбора файла
            textBox1.Text = openFileDialog1.FileName; //путь до выбранного файла отобразится в TextBox1

            if (textBox1.Text == "openFileDialog1") // если таблица не выбрана и диалоговое окно закрыто
            {
                textBox1.Text = ""; //очистка textBox1
                return; //завершение метода
            }
            
            excel = new Microsoft.Office.Interop.Excel.Application(); //создание объекта для работы с Excel
            wb = excel.Workbooks.Open(@textBox1.Text); //присвоение значений переменной
            ws = wb.Worksheets[1]; //присвоение значений переменной
            
            int RowsCount = ws.UsedRange.Rows.Count; //количество используемых строк

            object[,] Table = ws.get_Range("A1: A" + RowsCount).Value; //создается список городов

            if (Table == null || Convert.ToString(Table[1,1]) != "Города") //проверка на открытие нужной таблицы
            {
                MessageBox.Show("Выбрана неверная таблица"); //вывод ошибки
                textBox1.Clear(); //путь до файла в TextBox1 очищается
            }
            else
            {
                for (int i = 2; i < RowsCount+1; i++) //цикл для заполнения ComboBox
                {
                    comboBox1.Items.Add(Convert.ToString(Table[i, 1])); //заполнение списка ComboBox1 названиями городов
                    comboBox2.Items.Add(Convert.ToString(Table[i, 1])); //заполнение списка ComboBox2 названиями городов
                }
                comboBox1.Enabled = true; //открытие доступа ComboBox1
                comboBox2.Enabled = true; //открытие доступа ComboBox2
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (excel != null)
            {
                excel.Quit(); // закрытие Excel файла
                excel = null; // обнуление всех ссылок для корректного закрытия Excel файла
                wb = null;
                ws = null;
            }
            Close(); //закрывает форму
        }

        private void button4_Click(object sender, EventArgs e)
        {
            button4.Enabled = false; //закрытие доступа кнопки "Построить"
            listBox1.Items.Clear(); //очистка listBox

            Graph graph = new Graph(textBox1.Text, comboBox1.Text, comboBox2.Text); //создание объекта
            graph.PathCreate(); //вызов метода для построения кратчайшего пути
            
            foreach (var city in graph.ListOut) //цикл для вывода данных
            {
                listBox1.Items.Add(city); //все промежуточные точки пути запишутся в listBox1
            }

            listBox1.Visible = true; //открытие видимости listBox1
            label4.Text = Convert.ToString(graph.FinalPath); //присвоение lable4 итогового расстояния
            label3.Visible = true; //открытие видимости lable3
            label4.Visible = true; //открытие видимости lable4
            button3.Enabled = true; // открытие доступа кнопки "Очистить"

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) //изменение значения ComboBox1
        {
            if (comboBox1.Text != "" && comboBox2.Text != "") //если значения обоих ComboBox выбраны
            {
                button4.Enabled = true; //открытые доступа к кнопке "Построить"
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text != "" && comboBox2.Text != "") //если значения обоих ComboBox выбраны
            {
                button4.Enabled = true; //открытые доступа к кнопке "Построить"
            }
        }

        private void button3_Click(object sender, EventArgs e) //действия при нажатии на кнопку "Очистить"
        {
            button3.Enabled = false;
            listBox1.Items.Clear(); //очистка listBox
            label3.Visible = false; //скрытие Label3
            label4.Visible = false; //скрытие Label4
            button4.Enabled = true; //открытие доступа к кнопке "Построить"
        }
    }
    class Graph //создание класса
    {
        string path, combotext1, combotext2; //создание текстовых переменных
        public List<string> ListOut = new List<string>(); //лист для выавода итогового пути
        public int FinalPath; //переменная для вывода итогового расстояния
        
        public Graph (string a, string b, string c) //конструктор класса
        {
            path = a; //присвоение переменной Пути до файла
            combotext1 = b; //присвоение переменной Значения в ComboBox1
            combotext2 = c; //присвоение переменной Значения в ComboBox2
        }
        
        public void PathCreate() //метод для считывания исходных данных
        {
            int columnCount = Form1.ws.UsedRange.Columns.Count; //количество используемых столбцов
            string columnName = Form1.ws.Columns[columnCount].Address; //адрес последней используемой ячейки
            Regex reg = new Regex(@"(\$)(\w*):"); //регулярное выражение
            Match match = reg.Match(columnName); //представляет результаты из отдельного совпадения регулярного выражения
            string LastColumnName = match.Groups[2].Value; //имя последнего используемого столбца
            
            int count = columnCount-1; //количество столбцов с городами

            object[,] Table = Form1.ws.get_Range("A1:"+ LastColumnName + columnCount).Value; //создание объекта "матрицы смежности"
            Dictionary<string, Dictionary<string, int>> Points = new Dictionary<string, Dictionary<string, int>>(); //создание словаря словарей
            Dictionary<string, int> Communications = new Dictionary<string, int>(); //создание словаря

            for (int Row = 2; Row < count + 2; Row++) //цикл для внесения в программу связей между городами
            {
                Communications = new Dictionary<string, int>();

                for (int Column = 2; Column < count + 2; Column++) //цикл для поиска связей между городами
                {
                    if (Table[Row, Column] != null) //если значение ячейки не пустое
                        Communications[(string)Table[1, Column]] = Convert.ToInt32(Table[Row, Column]); //расстояние добавляется в соответствующую связь
                }
                Points[(string)Table[Row, 1]] = Communications; //заполнение словаря со связями
            }

            Table = null; //очистка объекта

            ShortestPath(combotext1, combotext2, Points); //вызов метода
        }


        public void ShortestPath(string start, string finish, Dictionary<string, Dictionary<string, int>> Points) //метод для построения кратчайшего пути
        {
            var previous = new Dictionary<string, string>(); // словарь для хранения предыдущих точек
            var distances = new Dictionary<string, int>(); //словарь для хранения расстояний между городами
            var nodes = new List<string>(); //лист для хранения узлов
            string FirstPoint = ""; // вспомогательная переменная для запоминания исходной точки

            List<string> path = null; //обнуление листа с итоговым путем
            int PathInt = 0; //обнуление итогового расстояния

            foreach (var point in Points) // заполнение словря расстояний из исходного города до всех городов
            {
                if (point.Key == start) //если это начальная точка
                    distances[point.Key] = 0; //расстояние до начальной точки всегда равно 0
                else
                    distances[point.Key] = int.MaxValue; //всем городам кроме начального присваивается максимальное значение

                nodes.Add(point.Key); //заполнение листа всеми городами
            }

            while (nodes.Count != 0) //пока не будут пройдены все вершины
            {
                nodes.Sort((x, y) => distances[x] - distances[y]); //сортировка всех городов в nodes по текущему расстоянию до них

                string smallest = nodes[0]; // smallest - непосещенный город с наименьшим расстоянием до него
                nodes.Remove(smallest); //удаление smallest из списка городов (город посещен)


                if (smallest == finish) //если дошли до конечной точки
                {
                    path = new List<string>(); //инициализация листа
                    while (previous.ContainsKey(smallest)) //цикл для заполнения листа с путем
                    {
                        path.Add(smallest); //добавление наименьшего в лист с путем
                        PathInt += Convert.ToInt32(Points[smallest][previous[smallest]]); //подсчет расстояния до конечного города
                        smallest = previous[smallest]; //присвоение нового значения переменной smallest
                    }
                    FirstPoint = smallest; //для технически верного отображения итогового пути
                    break; //выход из цикла
                }

                //if (distances[smallest] == int.MaxValue) //если расстояние до smallest не существует
                //{
                //    break;
                //}

                foreach (var neighbor in Points[smallest]) //цикл для поиска кратчайшего расстояния среди соседей текущей точки
                {
                    var alt = distances[smallest] + neighbor.Value; // расстояние до текущей + соседней
                    if (alt < distances[neighbor.Key]) //если найдено расстояние короче
                    {
                        distances[neighbor.Key] = alt; //присвоение нового кратчайшего расстояния
                        previous[neighbor.Key] = smallest; //запоминание пути
                    }
                }
            }

            FinalPath = PathInt; //присвоение итогового расстояния
            try //обработка исключений
            {
                if (path.Count != 0 || start == finish) //если не было ошибок ввода данных
                {
                    foreach (var city in path) //цикл для заполнения листа с промежуточными точками
                    {
                        ListOut.Add(city); //добавление промежуточных точек
                    }
                    ListOut.Add(FirstPoint); //добавление начального города
                }
                else ListOut.Add("Пути не существует"); //выведется в случае несуществования искомого пути
            }
            catch
            {
                ListOut.Add("Ошибка ввода данных"); //выведется в случае ошибки ввода данных
            }
        }
    }
}
