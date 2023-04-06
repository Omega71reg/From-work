public static void ОтчетПоКатегориямПВК(БланкДляВводаПланаПараметры бланкДляВводаПланаПараметры)
{
    /*if (бланкДляВводаПланаПараметры.ПустойБланк)
    {
        new Dictionary<string, object>().PrintToExcelTemplate("ПустойБланкДляВводаПлана");
        return;
    }*/
    List<ПозицияПланаВыпускаВКомплектахПоПроизводственнойВедомости> ПозПВК = new List<ПозицияПланаВыпускаВКомплектахПоПроизводственнойВедомости>();
    using (UnitOfWork uow = new UnitOfWork())
    {
        CriteriaOperator КритерийОтбора = CriteriaOperator.And
        (
            CriteriaOperator.Parse("ПроизводственнаяВедомость.ГруппаЗаказов Is Null"),
            CriteriaOperator.Or
            (
            new BinaryOperator("ПроизводственнаяВедомость", uow.GetObject(бланкДляВводаПланаПараметры.ПроизводственнаяВедомость)),
            new BinaryOperator("ПроизводственнаяВедомость.Тема", uow.GetObject(бланкДляВводаПланаПараметры.Тема)),
            new BinaryOperator("ПроизводственнаяВедомость.ОтветственныйОтПроизводственноДиспетчерскогоОтдела", uow.GetObject(бланкДляВводаПланаПараметры.ОтветственныйОтПдо))
            )
        );
        var tmpList1 = uow.GetObjects<ПозицияПланаВыпускаВКомплектахПоПроизводственнойВедомости>(КритерийОтбора);

        if (tmpList1 != null)
        {
            ПозПВК = tmpList1.ToList();
        }
    }
    var themes = ПозПВК.GroupBy(x => x.ПроизводственнаяВедомость?.Тема()).OrderBy(x => x.First().ПроизводственнаяВедомость?.Тема()?.Наименование).ToList();
    int themesCount = themes.Count;
    if (themesCount == 0)
    {
        return;
    }
    ExcelDocument excelDocument = new ExcelDocument { TemplateName = "БланкДляВводаПланаТЕСТ" };
    excelDocument.Init();
    excelDocument.CloneSheetAfter(1, excelDocument.SheetsCount, themesCount - 1);
    string[] monthes = DateTimeExtensions.GetMonthes();

    for (int i = 1; i <= themesCount; i++)
    {
        var theme = themes[i - 1];
        excelDocument.SetSheetName(i, String.IsNullOrEmpty(theme.First().ПроизводственнаяВедомость.Тема()?.Наименование) ? "Без темы" : theme.First().ПроизводственнаяВедомость.Тема()?.Наименование);

        excelDocument.SetSheetRangeByMarker(i, "#ТемаНаименованиеSheetNumber#", $"#ТемаНаименование{i}#");
        excelDocument.SetSheetRangeByMarker(i, "#ОтветственныйПдоНаименованиеSheetNumber#", $"#ОтветственныйПдоНаименование{i}#");
        excelDocument.SetSheetRangeByMarker(i, "#ТаблицаSheetNumber#", $"#Таблица{i}#");

        List<ExcelCellValueAndStyle[]> sheetData = new List<ExcelCellValueAndStyle[]>();
        var ведомости = ПозПВК?.GroupBy(x => x.ПроизводственнаяВедомость).OrderBy(x => x.First().ПроизводственнаяВедомость.Номер).ToList();

        foreach (var ведомость in ведомости.Where(x => x.Key.Тема() == theme.Key))
        {
            int curCount = sheetData.Count;
            sheetData.AddRange(ведомость.OrderBy(x => x.ПроизводственнаяЕдиница?.Код).ThenBy(x => x.Наименование).Select(x => new[] {

                        x.ПроизводственнаяВедомость?.Проект?.Номер.PackExcelData(),
                        x.ПроизводственнаяВедомость.Номер.PackExcelData(),
                        x.Наименование.PackExcelData(),
                        x.ПроизводственнаяЕдиница.КраткоеНаименование().PackExcelData(),
                        x.КоличествоКомплектовНаПроизводственнуюВедомость.PackExcelData(),
                        x.КоличествоКомплектовФакт.PackExcelData(),
                        "".PackExcelData(),
                        "".PackExcelData(),
                        "".PackExcelData(),
                    }));
            //sheetData[curCount][0].Value = ведомость.First().ПроизводственнаяВедомость.Номер;
            ExcelCellValueAndStyle.ClearRepeatedCellValues(sheetData, curCount, true, true, false, 1, 2, 3);
            sheetData.Add(new ExcelCellValueAndStyle[] { });
        }

        excelDocument.AddData($"ТемаНаименование{i}", $"Тема: {themes[i - 1].Key.Наименование}".PackExcelData());
        excelDocument.AddData($"ОтветственныйПдоНаименование{i}", $"Ответственный сотрудник ПДО: {themes[i - 1].Where(x => x.ПроизводственнаяВедомость?.ОтветственныйОтПроизводственноДиспетчерскогоОтдела() != null).GroupBy(x => x.ПроизводственнаяВедомость?.ОтветственныйОтПроизводственноДиспетчерскогоОтдела()).ToString(y => y.Key.Наименование)}".PackExcelData());
        excelDocument.AddData($"Таблица{i}", new ExcelTable { ReportRows = sheetData });
    }
    excelDocument.Print();
    excelDocument.Show();
}
    }