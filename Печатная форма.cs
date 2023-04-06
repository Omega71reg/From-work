using Galaktika.Kbp.Amm.Module.Models.ПланВыпускаОснастки;
using DevExpress.ExpressApp;
using Galaktika.Kbp.Amm.Module.Models.Planning;
using System.Collections.Generic;
using DevExpress.Xpo;
using DevExpress.Data.Filtering;
using Galaktika.Kbp.Amm.Module.Utils;
using System.Linq;
using Galaktika.Kbp.Amm.Module.Models;
using Galaktika.PRM.WOM.Module;
using Galaktika.PRM.ATP.Module;
using Galaktika.PRM.BOM.Module;
using System.Data;
using Galaktika.Kbp.Amm.Module.Models.Extensions;
using System;
using Galaktika.Core.Module;
using Galaktika.PRM.Cost.Module;
using Galaktika.Kbp.Amm.Module.Controllers.StateMachine;
using Galaktika.Core.CD.Module.RES;


static List<ЗаказПотребность> позицииЗПР = new List<ЗаказПотребность>();
static ЗаказНаСдачуПродукции ШапкаЗПР = null;


public string[] ToArray(БазовыйТехнологическийЭтап item)
{
    var result = new string[]
    {
                String.Format("{0}", item.ТехнологическоеОписание.ПредметПроизводства.Наименование),
                String.Format("{0:n0}",
                    (
                    item.ТехнологическоеОписание.КоличествоНаТехОтработку() > 0 ?
                    String.Format("{0:n0} + {1:n0}", item.ТехнологическоеОписание.КоличествоДсеНаПв(), item.ТехнологическоеОписание.КоличествоНаТехОтработку()) :
                    String.Format("{0:n0}", item.ТехнологическоеОписание.КоличествоДсеНаПв())
                    )),
                String.Format("{0}", item.Место.КраткоеНаименование()),
                String.Format("{0:n5}", item.ШтучноеВремя),
            String.Format("{0}", item.Статус())
    };
    return result;
}

public string[] ToArray2(Tuple<ПроизводственнаяЕдиница, decimal> item)
{

    var result = new string[]
    {
                String.Format("{0}",item.Item1.Наименование) ,
                String.Format("{0:n5}", item.Item2)
    };
    return result;
}
DataTable ConvertToDataTable1(List<БазовыйТехнологическийЭтап> list)
{
    DataTable table = new DataTable();

    table.Columns.Add("Содержание и объем работ", typeof(string));
    table.Columns.Add("Кол-во", typeof(string));
    table.Columns.Add("Цех", typeof(string));
    table.Columns.Add("Плановая трудоемкость", typeof(string));
    table.Columns.Add("Согласование ОТН УНиОТ", typeof(string));
    foreach (var item in list.OrderBy(row => row.ТехнологическоеОписание.ПредметПроизводства.Наименование))
    {
        table.Rows.Add(ToArray(item));
    }
    return table;
}

DataTable ConvertToDataTable2(List<БазовыйТехнологическийЭтап> list)
{
    DataTable table = new DataTable();

    table.Columns.Add("Цех", typeof(string));
    table.Columns.Add("Плановая трудоемкость", typeof(string));
    //var a = list.GroupBy(x => x.Место);
    /*var query = list.GroupBy(
        x => x.Место, (a,data) => new 
        {
            группацех = a.Наименование,
            суммавремя = data.Sum(x => x.ШтучноеВремя)
        }
        ).ToList();*/
    foreach (var item in list.GroupBy(x => x.Место))
    //.OrderBy(row => row.ТехнологическоеОписание.ПредметПроизводства.Наименование))
    {
        table.Rows.Add(ToArray2(new Tuple<ПроизводственнаяЕдиница, decimal>(item.Key, item.Sum(x => x.ШтучноеВремя))));
    }
    return table;
}
string ПВ1 = "", заказы = "";
decimal a = 0;

private void ЗаполнениеТаблицы_BeforePrint(object sender, System.Drawing.Printing.PrintEventArgs e)
{
    ШапкаЗПР = GetCurrentRow() as ЗаказНаСдачуПродукции;
    //System.Windows.Forms.MessageBox.Show((ШапкаЗПР.Oid).ToString(), "Позиции");

    if (ШапкаЗПР.ГруппаЗаказов == null)
    {
        System.Windows.Forms.MessageBox.Show(String.Format("Это не ЗПР"), "Oops...", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
        this.StopPageBuilding();
        return;
    }

    using (UnitOfWork uow = new UnitOfWork())
    {
        CriteriaOperator критерия = CriteriaOperator.And
            (
            new BinaryOperator("ЗаказНаСдачуП.Oid", ШапкаЗПР.Oid)
            );

        var tmpList1 = uow.GetObjects<ЗаказПотребность>(критерия);

        if (tmpList1 != null)
        {
            позицииЗПР = tmpList1.Where(x => x.ТехОписание != null).ToList();
            //System.Windows.Forms.MessageBox.Show((позицииЗПР.Count).ToString(),"Позиции")
        }

        ПВ1 = string.Join(", ", позицииЗПР.Where(item => item.ЗаказГП != null).GroupBy(item => item.ЗаказГП).Select(item => item.FirstOrDefault().ЗаказГП.Номер));
        ПВ.Text = ПВ1.ToString();

        заказы = позицииЗПР.Count > 0 ?
              string.Join(", ", позицииЗПР.GroupBy(item => item.Проект).Select(item => item.FirstOrDefault().Проект.Номер)) :
              ШапкаЗПР.Проект.Номер;
        Заказ.Text = заказы;


        //System.Windows.Forms.MessageBox.Show(str);
        List<БазовыйТехнологическийЭтап> цехозаходы2 = позицииЗПР.SelectMany(x => x.ТехОписание.МаршрутБезИсключенныхВидовРабот()).ToList();

        a = цехозаходы2.Sum(x => x.ШтучноеВремя);
        label28.Text = String.Format("{0:n5}", a);

        DataTable table1 = ConvertToDataTable1(цехозаходы2);
        DetailReport.DataSource = table1;

        DataTable table2 = ConvertToDataTable2(цехозаходы2);
        DetailReport3.DataSource = table2;
    }
    tableCell5.DataBindings.Add("Text", null, "Содержание и объем работ");
    tableCell11.DataBindings.Add("Text", null, "Кол-во");
    tableCell6.DataBindings.Add("Text", null, "Цех");
    tableCell7.DataBindings.Add("Text", null, "Плановая трудоемкость");
    tableCell8.DataBindings.Add("Text", null, "Согласование ОТН УНиОТ");

    tableCell10.DataBindings.Add("Text", null, "Цех");
    tableCell12.DataBindings.Add("Text", null, "Плановая трудоемкость");
}

private void ЗаполнениеШапки_BeforePrint(object sender, System.Drawing.Printing.PrintEventArgs e)
{
    ШапкаЗПР = GetCurrentRow() as ЗаказНаСдачуПродукции;

    if (ШапкаЗПР.ГруппаЗаказов == null)
    {
        System.Windows.Forms.MessageBox.Show(String.Format("Это не ЗПР"), "Oops...", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
        this.StopPageBuilding();
        return;
    }


    if (ШапкаЗПР != null)
    {
        подразделениеАвтора.Text = ШапкаЗПР.CreateEmployeeCalc().Сектор().Наименование;

        ДатаИсполнения.Text = String.Format("от {0} № {1}",
            ШапкаЗПР.ДатаДокумента,
            ШапкаЗПР.Номер);

        /*ВходящийНомер.Text = String.Format("Входящий номер в производстве {0}",
            ШапкаЗПР.GetMemberValue("ВходящийНомерВПроизводстве"));*/
    }
}


private void ЗаполнениеПодвал_BeforePrint(object sender, System.Drawing.Printing.PrintEventArgs e)
{
    ШапкаЗПР = GetCurrentRow() as ЗаказНаСдачуПродукции;

    if (ШапкаЗПР.ГруппаЗаказов == null)
    {
        System.Windows.Forms.MessageBox.Show(String.Format("Это не ЗПР"), "Oops...", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
        this.StopPageBuilding();
        return;
    }


    if (ШапкаЗПР != null)
    {
        label31.Text = ШапкаЗПР.CreateEmployeeCalc().Сектор().Руководитель == null ? "" : ШапкаЗПР.CreateEmployeeCalc().Сектор().Руководитель.GetMemberValue("Наименование").ToString();

        /*label31.Text = String.Format("Начальник {0}",
            ШапкаЗПР.CreateEmployeeCalc().Сектор().Наименование);*/

        label30.Text = String.Format("Исп. {0}",
                ШапкаЗПР.CreateEmployeeCalc().Наименование);

        /*ВходящийНомер.Text = String.Format("Входящий номер в производстве {0}",
            ШапкаЗПР.GetMemberValue("ВходящийНомерВПроизводстве"));*/

        string name2 = ШапкаЗПР.CreateEmployeeCalc().Сектор().Наименование as string;
        string result = "Начальник " + name2.Replace("ОТДЕЛЕНИЕ", "отделения");
        label29.Text = result;

        if (ШапкаЗПР.СтатусЭлектронногоСогласования().Code() > 405)
        {
            label10.Visible = true;
            if (ШапкаЗПР.ИсторияСогласованияList().Where(x => x.СтатусКонечный.Code() == 406).Count() > 0)
                label11.Text = ШапкаЗПР.ИсторияСогласованияList().Where(x => x.СтатусКонечный.Code() == 406).OrderByDescending(x => x.Date).First().Date.ToString();
        }


        //            System.Windows.Forms.MessageBox.Show((ШапкаЗПР.ИсторияСогласованияList().Count).ToString(), "колво");*/


        /*bool a = true;
        if (ШапкаЗПР.СтатусЭлектронногоСогласования().As<СостояниеЗапроса>() == СостояниеЗапроса.ВРаботеЦехов)
        label15.Visible = a;  */
    }
}

private void СогласованиеЗапроса_BeforePrint(object sender, System.Drawing.Printing.PrintEventArgs e)
{

    if (ШапкаЗПР.ГруппаЗаказов == null)
    {
        System.Windows.Forms.MessageBox.Show(String.Format("Это не ЗПР"), "Oops...", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
        this.StopPageBuilding();
        return;
    }


    if (ШапкаЗПР.СтатусЭлектронногоСогласования().Code() == 460)
    {

        if (ШапкаЗПР.ИсторияСогласованияList().Where(x => x.СтатусКонечный.Code() == 408).Count() > 0)
        {
            label40.Text = ШапкаЗПР.ИсторияСогласованияList().Where(x => x.СтатусКонечный.Code() == 408).OrderByDescending(x => x.Date).First().Сотрудник.Наименование.ToString();
            label35.Text = ШапкаЗПР.ИсторияСогласованияList().Where(x => x.СтатусКонечный.Code() == 408).OrderByDescending(x => x.Date).First().Date.ToString();
        }

        if (ШапкаЗПР.ИсторияСогласованияList().Where(x => x.СтатусКонечный.Code() == 430).Count() > 0)
        {
            label39.Text = ШапкаЗПР.ИсторияСогласованияList().Where(x => x.СтатусКонечный.Code() == 430).OrderByDescending(x => x.Date).First().Сотрудник.Наименование.ToString();
            label34.Text = ШапкаЗПР.ИсторияСогласованияList().Where(x => x.СтатусКонечный.Code() == 430).OrderByDescending(x => x.Date).First().Date.ToString();
        }

        if (ШапкаЗПР.ИсторияСогласованияList().Where(x => x.СтатусНачальный.Code() == 445).Count() > 0)
        {
            label41.Text = ШапкаЗПР.ИсторияСогласованияList().Where(x => x.СтатусНачальный.Code() == 445).OrderByDescending(x => x.Date).First().Сотрудник.Наименование.ToString();
            label33.Text = ШапкаЗПР.ИсторияСогласованияList().Where(x => x.СтатусНачальный.Code() == 445).OrderByDescending(x => x.Date).First().Date.ToString();
        }

        if (ШапкаЗПР.ИсторияСогласованияList().Where(x => x.СтатусНачальный.Code() == 450).Count() > 0)
        {
            label43.Text = ШапкаЗПР.ИсторияСогласованияList().Where(x => x.СтатусНачальный.Code() == 450).OrderByDescending(x => x.Date).First().Сотрудник.Наименование.ToString();
            label32.Text = ШапкаЗПР.ИсторияСогласованияList().Where(x => x.СтатусНачальный.Code() == 450).OrderByDescending(x => x.Date).First().Date.ToString();
        }
    }
    else ReportFooter1.Visible = false;


}

