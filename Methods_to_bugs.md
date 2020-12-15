# Methods with errors

## Вопросы по API методам

- Как работает History Point

## Document Editor

### Text document API/Api/CreateBlockLvlSdt

Выдает ошибку при попытке вызвать метод.
Следовательно, не работают также все методы класса ApiBlockLvlSdt.

Text document API/ApiBlockLvlSdt/GetAlias
Text document API/ApiBlockLvlSdt/GetClassType
Text document API/ApiBlockLvlSdt/GetContent
Text document API/ApiBlockLvlSdt/GetLabel
Text document API/ApiBlockLvlSdt/GetLock
Text document API/ApiBlockLvlSdt/GetTag
Text document API/ApiBlockLvlSdt/SetAlias
Text document API/ApiBlockLvlSdt/SetLabel
Text document API/ApiBlockLvlSdt/SetLock
Text document API/ApiBlockLvlSdt/SetTag

### Text document API/Api/CreateBullet

Ошибок не выдает, но bullet перед строкой не выводится. `В sdkjs указано, что этот метод только для CSE и CPE`. Следовательно, неправильно работают и методы:

builder.CreateFile("docx");
var oDocument = Api.GetDocument();
var oParagraph = oDocument.GetElement(0);
var oBullet = Api.CreateBullet("-");
oParagraph.SetBullet(oBullet);
oParagraph.AddText("This is an example of the bulleted paragraph.");
builder.SaveFile("docx", "CreateBullet.docx");
builder.CloseFile();

(Version: 5.6.3 (build:2))

>Желательно после описания проблемы скинуть скрипт который не работает и версию редактора на котором тестировался.

Text document API/Api/CreateNumbering
Text document API/ApiParagraph/SetBullet
Text document API/ApiParaPr/SetBullet

>Косяк в документации
    Для CDE есть CreateNumbering, нужно работать через него
    SetBullet - это для презентаций и таблиц

```js
oDocument = Api.GetDocument();
oParagraph = Api.CreateParagraph();
oNumbering = oDocument.CreateNumbering(); // Надо применять метод к документу, тогда создается объект function(){return"numbering"} Если применять к Api создаем «bullet»
alert(oNumbering.GetClassType);
oParagraph.AddText("Text");
oNumLvl = oNumbering.GetLevel(1); // function(){return"numberingLevel"}
alert(oNumLvl.GetClassType);
oNumLvl.SetTemplateType("1)");
oParagraph.SetNumbering(oNumLvl);
oDocument.Push(oParagraph);
```

### Text document API/ApiDocument/SearchAndReplace

Что мы должны передавать в параметр oProperties? Paragraph? Но с ним у меня не работает.
Применяя метод SearchAndReplace мы передаем в параметрах json объект, в котором указываем текст, который будет найден в документе и текст которым мы заменим "выделенное".

```js
(function()
{
    var oDocument = Api.GetDocument();
    oDocument.SearchAndReplace({"searchString": "qweqwe", "replaceString": "lalala"});
})();
```

### Text document API/ApiUnsupported/GetClassType

Что это за класс и как его можно отразить в самом документе?
это просто объект для ситуаций, когда у нас нет api для класса, который может вернуться откуда-нибудь
Я вроде поняла, но как это отразить в примере?

builder.CreateFile("docx");
var Document = Api.GetDocument();
var Paragraph = Document.GetElement(0);
var Unsupported = ();
var ClassType = Unsupported.GetClassType();
Paragraph.AddText("Class Type = " + ClassType);
builder.SaveFile("docx", "GetClassType.docx");
builder.CloseFile();

(Version: 5.6.3 (build:2))

Ответ от разрабов:
    допустим метод GetElement(nPos) у параграфа или другого объекта, который может в себе другие объекты содержать. Вдруг возвращается объект, для которого у нас api не написан, т.е. проверка не знает что пришло, что делать в таком случае? Отдать пользователю объект без api?

>Думаю что нет особого смысла писать пример на данный метод тк это заглушка просто на такие ситуации

### Text document API/ApiTable/Copy

У меня метод почему-то не работает (выдает ошибку).

>Ошибка на шаге копирования, отправил разрабам

builder.CreateFile("docx");
var oDocument = Api.GetDocument();
var oTable1 = Api.CreateTable(2, 3);
oTable1.SetWidth("percent", 100);
var oTable2 = oTable1.Copy();
oDocument.Push(oTable1);
oDocument.Push(oTable2);
builder.SaveFile("docx", "Copy.docx");
builder.CloseFile();    (Version: 5.6.3 (build:2)

### Text document API/ApiDocument/GetAllTablesOnPage

Этот код работает у меня как-то странно. Когда я запускаю его в первый раз, просто появляются две таблицы, но ряд не удаляется. Когда я запускаю его сразу после этого еще один раз, то появляются еще две таблицы и ряд у одной из них уже удаляется (как и должен). Т.е. метод вроде работает, но я не понимаю, почему только со второго раза.

```js
builder.CreateFile("docx");
var oDocument = Api.GetDocument();
var oTableStyle = oDocument.CreateStyle("CustomTableStyle", "table");
oTableStyle.SetBasedOn(oDocument.GetStyle("Bordered - Accent 5"));
var oTable1 = Api.CreateTable(3, 3);
oTable1.SetWidth("percent", 100);
oTable1.SetStyle(oTableStyle);
oDocument.Push(oTable1);
var oTable2 = Api.CreateTable(2, 2);
oTable2.SetWidth("percent", 100);
oTable2.SetStyle(oTableStyle);
oDocument.Push(oTable2);
var oTables = oDocument.GetAllTablesOnPage(0);
oCell = oTables[0].GetRow(1).GetCell(0);
oTables[0].RemoveRow(oCell);
builder.SaveFile("docx", "GetAllTablesOnPage.docx");
builder.CloseFile();
```

>Во-первых данный метод можно в два раза упростить

```js
    var oDocument = Api.GetDocument();
    oTableStyle = oDocument.GetStyle("Bordered - Accent 5")
    var oTable = Api.CreateTable(3, 3);
    oTable.SetWidth("percent", 50);
    oTable.SetStyle(oTableStyle);
    oDocument.Push(oTable);
    var arrTables = oDocument.GetAllTablesOnPage(0);
    oRow_1 = arrTables[0].GetRow(0);
    oRow_1.Remove();
```

> Во-вторых написал баг на Никиту Хромова

<https://bugzilla.onlyoffice.com/show_bug.cgi?id=48023>

Version: 6.0.0 (build:105)

### Text document API/ApiChart/GetPrevChart

Ошибка при вызове метода.

```js
builder.CreateFile("docx");
var oDocument = Api.GetDocument();
var oParagraph = oDocument.GetElement(0);
var oChart = Api.CreateChart("bar3D", [
  [200, 240, 280],
  [250, 260, 280]
], ["Projected Revenue", "Estimated Costs"], [2014, 2015, 2016], 4051300, 2347595, 24);
oChart.SetVerAxisTitle("USD In Hundred Thousands", 10);
oChart.SetHorAxisTitle("Year", 11);
oChart.SetLegendPos("bottom");
oChart.SetShowDataLabels(false, false, true, false);
oChart.SetTitle("Financial Overview", 13);
oParagraph.AddDrawing(oChart);
var oCopyChart = oChart.Copy();
oParagraph.AddDrawing(oCopyChart);
var oPrevChart = oCopyChart.GetPrevChart();
var oStroke = Api.CreateStroke(1 * 150, Api.CreateSolidFill(Api.CreateRGBColor(155, 64, 1)));
oPrevChart.SetMinorHorizontalGridlines(oStroke);
builder.SaveFile("docx", "GetPrevChart.docx");
builder.CloseFile();
```

>Написал баг на Никиту Хромова

<https://bugzilla.onlyoffice.com/show_bug.cgi?id=48027>

Version: 6.0.0 (build:105)

### Text document API/ApiParagraph/GetAllShapes

`Метод срабатывает, но почему-то на строке oDrawings[1].Fill(oFill); возникает ошибка. Хотя точно такой же пример, но для ApiDocument работает.`

```js
builder.CreateFile("docx");
var oDocument = Api.GetDocument();
var oParagraph = oDocument.GetElement(0);
var oGs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 224, 204), 0);
var oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 164, 101), 100000);
var oFill = Api.CreateLinearGradientFill([oGs1, oGs2], 5400000);
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
var oDrawing1 = Api.CreateShape("rect", 3212465, 963295, oFill, oStroke);
oParagraph.AddDrawing(oDrawing1);
var oDrawing2 = Api.CreateShape("wave", 3212465, 963295, oFill, oStroke);
oParagraph.AddDrawing(oDrawing2);
var oDrawings = oParagraph.GetAllShapes();
oFill = Api.CreateSolidFill(Api.CreateRGBColor(61, 74, 107));
oDrawings[1].Fill(oFill);
builder.SaveFile("docx", "GetAllShapes.docx");
builder.CloseFile();
```

Version: 6.0.0 (build:105)

### Text document API/ApiParagraph/GetAllCharts

`Та же проблема, что и с предыдущим методом.`

```js
builder.CreateFile("docx");
var oDocument = Api.GetDocument();
var oParagraph = oDocument.GetElement(0);
var oChart1 = Api.CreateChart("bar3D", [
  [200, 240, 280],
  [250, 260, 280]
], ["Projected Revenue", "Estimated Costs"], [2014, 2015, 2016], 4051300, 2347595, 24);
oParagraph.AddDrawing(oChart1);
var oChart2 = Api.CreateChart("bar3D", [
  [200, 240, 280],
  [250, 260, 280]
], ["Projected Revenue", "Estimated Costs"], [2014, 2015, 2016], 4051300, 2347595, 24);
oChart2.SetVerAxisTitle("USD In Hundred Thousands", 10);
oChart2.SetHorAxisTitle("Year", 11);
oChart2.SetLegendPos("bottom");
oChart2.SetShowDataLabels(false, false, true, false);
oChart2.SetTitle("Financial Overview", 13);
oParagraph.AddDrawing(oChart2);
var oCharts = oParagraph.GetAllCharts();
oStroke = Api.CreateStroke(1 * 150, Api.CreateSolidFill(Api.CreateRGBColor(155, 64, 1)));
oCharts[1].SetMinorHorizontalGridhlines(oStroke);
builder.SaveFile("docx", "GetAllCarts.docx");
builder.CloseFile();
```

Version: 6.0.0 (build:105)

### Text document API/ApiParagraph/GetParentTable и GetParentTableCell

`Ошибка при добавлении параграфа в таблицу или ячейку. Проблема с методом AddElement. Тоже самое и для класса ApiRun.`

```js
builder.CreateFile("docx");
oDocument = Api.GetDocument();
oTableStyle = oDocument.CreateStyle("CustomTableStyle", "table");
oTableStyle.SetBasedOn(oDocument.GetStyle("Bordered - Accent 5"));
oTable = Api.CreateTable(3, 3);
oTable.SetWidth("percent", 100);
oTable.SetStyle(oTableStyle);
oDocument.Push(oTable);
oParagraph = Api.CreateParagraph();
oParagraph.AddText("This is just a sample text.");
oTable.AddElement(0, oParagraph);
oParentTable = oParagraph.GetParentTable();
oCell = oParentTable.GetRow(2).GetCell(0);
oParentTable.RemoveRow(oCell);
builder.SaveFile("docx", "GetParentTable.docx");
builder.CloseFile();
```

Version: 6.0.0 (build:105)

### Text document API/ApiHyperlink/GetDisplayedText

`Ошибка при вызове метода.`

```js
builder.CreateFile("docx");
oDocument = Api.GetDocument();
oParagraph = oDocument.GetElement(0);
oRun = Api.CreateRun();
oRun.AddText("ONLYOFFICE Document Builder");
oParagraph.AddElement(oRun);
oHyperlink = oParagraph.AddHyperlink("http://api.teamlab.info/docbuilder/basic");
oHyperlink.SetDisplayedText("Api ONLYOFFICE DocBuilder");
oDisplayedText = oHyperlink.GetDisplayedText();
oParagraph = Api.CreateParagraph();
oParagraph.AddText("Displayed text: " + oDisplayedText);
oDocument.Push(oParagraph);
builder.SaveFile("docx", "GetDisplayedText.docx");
builder.CloseFile();
```

Version: 6.0.0 (build:105)

### Text document API/ApiHyperlink/GetElement и GetElementsCount

`Метод GetElement() ничего не выводит при индексе 0, первый run выводит при индексе 1, а второй при индексе 2. При этом метод GetElementsCount выдает значение 4, хотя у нас в параграфе всего 2 run.`

```js
builder.CreateFile("docx");
oDocument = Api.GetDocument();
oParagraph = oDocument.GetElement(0);
oParagraph.AddText("Api Document Builder.");
oRun = Api.CreateRun();
oRun.AddText(" ONLYOFFICE for developers");
oParagraph.AddElement(oRun);
oHyperlink = oParagraph.AddHyperlink("http://api.teamlab.info/docbuilder/basic");
oElement = oHyperlink.GetElement(1);
oParagraph = Api.CreateParagraph();
oParagraph.AddElement(oElement);
oDocument.Push(oParagraph);
builder.SaveFile("docx", "GetElement.docx");
builder.CloseFile();
```

```js
builder.CreateFile("docx");
oDocument = Api.GetDocument();
oParagraph = oDocument.GetElement(0);
oParagraph.AddText("Api Document Builder.");
oRun = Api.CreateRun();
oRun.AddText(" ONLYOFFICE for developers");
oParagraph.AddElement(oRun);
oHyperlink = oParagraph.AddHyperlink("http://api.teamlab.info/docbuilder/basic");
oElementsCount = oHyperlink.GetElementsCount();
oParagraph = Api.CreateParagraph();
oParagraph.AddText("Number of elements in hyperlink: " + oElementsCount);
oDocument.Push(oParagraph);
builder.SaveFile("docx", "GetElementsCount.docx");
builder.CloseFile();
```

Version: 6.0.0 (build:105)

### Text document API/ApiHyperlink/SetDefaultStyle

`Что имеется в виду под стилем гиперссылки? Если стиль отображаемого текста, но метод у меня не срабатывает.`

```js
builder.CreateFile("docx");
oDocument = Api.GetDocument();
oParagraph = oDocument.GetElement(0);
oParagraph.AddText("Api Document Builder");
oHyperlink = oParagraph.AddHyperlink("http://api.teamlab.info/docbuilder/basic");
oText = oHyperlink.GetLinkedText();
oParagraph.SetFontSize(16);
oHyperlink.SetDefaultStyle();
builder.SaveFile("docx", "SetDefaultStyle .docx");
builder.CloseFile();
```

Version: 6.0.0 (build:105)

### Text document API/ApiRun/GetParentContentControl

```js
При вызове метода - false.
builder.CreateFile("docx");
oDocument = Api.GetDocument();
oParagraph = oDocument.GetElement(0);
oInlineLvlSdt = Api.CreateInlineLvlSdt();
oParagraph.AddInlineLvlSdt(oInlineLvlSdt);
oRun = Api.CreateRun();
oRun.AddText("This is an inline text content control.");
oInlineLvlSdt.AddElement(oRun, 0);
oContentControl = oRun.GetParentContentControl();
oClassType = oContentControl.GetClassType();
oParagraph = Api.CreateParagraph();
oParagraph.AddText("Class type: " + oClassType);
oDocument.Push(oParagraph);
builder.SaveFile("docx", "GetParentContentControl.docx");
builder.CloseFile();
```

Version: 6.0.0 (build:105)

### Text document API/ApiSection/GetNext и GetPrevious

```js
Методы не работают.
builder.CreateFile("docx");
oDocument = Api.GetDocument();
oParagraph = oDocument.GetElement(0);
oParagraph.AddText("This is a new paragraph.");
oParagraph.AddLineBreak();
oParagraph.AddText("Scroll down to see the new section.");
oSection1 = oDocument.CreateSection(oParagraph);
oParagraph = Api.CreateParagraph();
oParagraph.AddText("This is a paragraph in a new section");
oDocument.Push(oParagraph);
oSection2 = oDocument.CreateSection(oParagraph);
oParagraph = Api.CreateParagraph();
oParagraph.AddText("This is a paragraph in a new section");
oDocument.Push(oParagraph);
oPreviousSection = oSection2.GetPrevious();
oHeader = oPreviousSection.GetHeader("default", true);
oParagraph = oHeader.GetElement(0);
oParagraph.AddText("This is a page header");
builder.SaveFile("docx", "GetPrevious.docx");
builder.CloseFile();
```

Version: 6.0.0 (build:105)

```js
builder.CreateFile("docx");
oDocument = Api.GetDocument();
oParagraph = oDocument.GetElement(0);
oParagraph.AddText("This is a new paragraph.");
oParagraph.AddLineBreak();
oParagraph.AddText("Scroll down to see the new section.");
oSection1 = oDocument.CreateSection(oParagraph);
oParagraph = Api.CreateParagraph();
oParagraph.AddText("This is a paragraph in a new section");
oDocument.Push(oParagraph);
oNextSection = oSection1.GetNext();
oHeader = oNextSection.GetHeader("default", true);
oParagraph = oHeader.GetElement(0);
oParagraph.AddText("This is a page header");
builder.SaveFile("docx", "GetNext.docx");
builder.CloseFile();
```

Version: 6.0.0 (build:105)

### Text document API/ApiTable/Split

Если ставить количество столбцов, на которые нужно разбить ячейку, не больше 1, то все работает. Но при установке параметра nCols больше единицы, у меня вообще все зависает и приходится перезагружать документ. При повторном открытии этого документа появляется таблица, но совсем не такая, какая должна быть. То же самое и для ApiTableCell.

```js
builder.CreateFile("docx");
oDocument = Api.GetDocument();
oTableStyle = oDocument.CreateStyle("CustomTableStyle", "table");
oTableStyle.SetBasedOn(oDocument.GetStyle("Bordered - Accent 5"));
oTable = Api.CreateTable(3, 3);
oTable.SetWidth("percent", 100);
oTable.SetStyle(oTableStyle);
oDocument.Push(oTable);
oCell = oTable.GetCell(0, 0);
oTable.Split(oCell, 2, 1);
builder.SaveFile("docx", "Split.docx");
builder.CloseFile();
```

Version: 6.0.0 (build:105)

### Text document API/ApiTable/AddElement

`Как я уже писала выше, этот метод у меня не работает. Такая же проблема с этим методом и для ApiTableCell.`

```js
builder.CreateFile("docx");
oDocument = Api.GetDocument();
oTableStyle = oDocument.CreateStyle("CustomTableStyle", "table");
oTableStyle.SetBasedOn(oDocument.GetStyle("Bordered - Accent 5"));
oTable = Api.CreateTable(3, 3);
oTable.SetWidth("percent", 100);
oTable.SetStyle(oTableStyle);
oDocument.Push(oTable);
oParagraph = Api.CreateParagraph();
oParagraph.AddText("This is just a sample text in the first cell.");
oTable.AddElement(0, oParagraph);
builder.SaveFile("docx", "AddElement.docx");
builder.CloseFile();
```

Version: 6.0.0 (build:105)

### Text document API/ApiTable/GetNext

Не работает. Возвращает null, хотя таблица не последняя. То же и с GetPrevious.

```js
builder.CreateFile("docx");
oDocument = Api.GetDocument();
oTableStyle = oDocument.CreateStyle("CustomTableStyle", "table");
oTableStyle.SetBasedOn(oDocument.GetStyle("Bordered - Accent 5"));
oTable1 = Api.CreateTable(3, 3);
oTable1.SetWidth("percent", 100);
oTable1.SetStyle(oTableStyle);
oDocument.Push(oTable1);
oTable1.GetCell(0, 0).GetContent().GetElement(0).AddText("Table 1");
oTable2 = Api.CreateTable(3, 3);
oTable2.SetWidth("percent", 100);
oTable2.SetStyle(oTableStyle);
oDocument.Push(oTable2);
oTable2.GetCell(0, 0).GetContent().GetElement(0).AddText("Table 2");
oNextTable = oTable1.GetNext();
oNextTable.SetTableBorderTop("single", 32, 0, 0, 0, 255);
builder.SaveFile("docx", "GetNext.docx");
builder.CloseFile();
```

```js
builder.CreateFile("docx");
oDocument = Api.GetDocument();
oTableStyle = oDocument.CreateStyle("CustomTableStyle", "table");
oTableStyle.SetBasedOn(oDocument.GetStyle("Bordered - Accent 5"));
oTable1 = Api.CreateTable(3, 3);
oTable1.SetWidth("percent", 100);
oTable1.SetStyle(oTableStyle);
oDocument.Push(oTable1);
oTable1.GetCell(0, 0).GetContent().GetElement(0).AddText("Table 1");
oTable2 = Api.CreateTable(3, 3);
oTable2.SetWidth("percent", 100);
oTable2.SetStyle(oTableStyle);
oDocument.Push(oTable2);
oTable2.GetCell(0, 0).GetContent().GetElement(0).AddText("Table 2");
oPreviousTable = oTable2.GetPrevious();
oPreviousTable.SetTableBorderTop("single", 32, 0, 0, 0, 255);
builder.SaveFile("docx", "GetPrevious.docx");
builder.CloseFile();
```

Version: 6.0.0 (build:105)

### Text document API/ApiTable/Delete

`Возвращает true, но таблица при этом не удаляется.`

```js
builder.CreateFile("docx");
oDocument = Api.GetDocument();
oTableStyle = oDocument.CreateStyle("CustomTableStyle", "table");
oTableStyle.SetBasedOn(oDocument.GetStyle("Bordered - Accent 5"));
oTable = Api.CreateTable(3, 3);
oTable.SetWidth("percent", 100);
oTable.SetStyle(oTableStyle);
oDocument.Push(oTable);
oTable.Delete();
oParagraph = Api.CreateParagraph();
oParagraph.AddText("The table was removed from the document.");
oDocument.Push(oParagraph);
builder.SaveFile("docx", "Delete.docx");
builder.CloseFile();
```

Version: 6.0.0 (build:105)

### Text document API/Api

`Ошибки при вызове методов`

1. CreateRange
2. CreateHyperlink
3. AddComment
4. GetFirstRunInArray
5. GetLastRunInArray

### Text document API/ApiBlockLvlSdt/Delete

`Блок не удаляется.`

```js
builder.CreateFile("docx");
oDocument = Api.GetDocument();
oBlockLvlSdt = Api.CreateBlockLvlSdt();
oBlockLvlSdt.AddText("This is a block text content control.");
oDocument.AddElement(0, oBlockLvlSdt);
oBlockLvlSdt.Delete(false);
oParagraph = oDocument.GetElement(0);
oParagraph.AddText("The block text content control was removed from the document.")
builder.SaveFile("docx", "Delete.docx");
builder.CloseFile();
```

### Text document API/ApiBlockLvlSdt/GetAllContentControls

`Метод не работает.`

```js
builder.CreateFile("docx");
oDocument = Api.GetDocument();
oBlockLvlSdt = Api.CreateBlockLvlSdt();
oBlockLvlSdt1 = Api.CreateBlockLvlSdt();
oBlockLvlSdt1.AddText("This is the first block text content control.");
oBlockLvlSdt2 = Api.CreateBlockLvlSdt();
oBlockLvlSdt2.AddText("This is the second block text content control.");
oBlockLvlSdt.AddElement(oBlockLvlSdt1, 0);
oBlockLvlSdt.AddElement(oBlockLvlSdt2, 1);
oDocument.AddElement(0, oBlockLvlSdt);
aContentControls = oBlockLvlSdt.GetAllContentControls();
oParagraph = Api.CreateParagraph();
oParagraph.AddText("Class type of the first element in array: " + aContentControls[0].GetClassType());
builder.SaveFile("docx", "GetAllContentControls.docx");
builder.CloseFile();
```

### Text document API/ApiBlockLvlSdt/GetAllDrawingObjects

`Метод не работает.`

```js
builder.CreateFile("docx");
oDocument = Api.GetDocument();
oBlockLvlSdt = Api.CreateBlockLvlSdt();
oGs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 224, 204), 0);
oGs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 164, 101), 100000);
oFill = Api.CreateLinearGradientFill([oGs1, oGs2], 5400000);
oStroke = Api.CreateStroke(0, Api.CreateNoFill());
oDrawing1 = Api.CreateShape("rect", 3212465, 963295, oFill, oStroke);
oBlockLvlSdt.AddElement(oDrawing1, 0);
oDrawing2 = Api.CreateChart("bar3D", [
    [200, 240, 280],
    [250, 260, 280]
], ["Projected Revenue", "Estimated Costs"], [2014, 2015, 2016], 4051300, 2347595, 24);
oDrawing2.SetVerAxisTitle("USD In Hundred Thousands", 10);
oDrawing2.SetHorAxisTitle("Year", 11);
oDrawing2.SetLegendPos("bottom");
oDrawing2.SetShowDataLabels(false, false, true, false);
oDrawing2.SetTitle("Financial Overview", 13);
oBlockLvlSdt.AddElement(oDrawing2, 1);
oDocument.AddElement(0, oBlockLvlSdt);
aDrawingObjects = oBlockLvlSdt.GetAllDrawingObjects();
oFill = Api.CreateSolidFill(Api.CreateRGBColor(61, 74, 107));
aDrawingObjects[0].Fill(oFill);
builder.SaveFile("docx", "GetAllDrawingObjects.docx");
builder.CloseFile();
```

### Text document API/ApiBlockLvlSdt/GetAllTablesOnPage

`Метод не работает.`

```js
builder.CreateFile("docx");
oDocument = Api.GetDocument();
oBlockLvlSdt = Api.CreateBlockLvlSdt();
oTableStyle = oDocument.CreateStyle("CustomTableStyle", "table");
oTableStyle.SetBasedOn(oDocument.GetStyle("Bordered - Accent 5"));
oTable1 = Api.CreateTable(3, 3);
oTable1.SetWidth("percent", 100);
oTable1.SetStyle(oTableStyle);
oBlockLvlSdt.AddElement(oTable1, 0);
oTable2 = Api.CreateTable(2, 2);
oTable2.SetWidth("percent", 100);
oTable2.SetStyle(oTableStyle);
oBlockLvlSdt.AddElement(oTable2, 1);
oDocument.AddElement(0, oBlockLvlSdt);
aTables = oBlockLvlSdt.GetAllTablesOnPage();
oCell = aTables[0].GetRow(1).GetCell(0);
aTables[0].RemoveRow(oCell);
builder.SaveFile("docx", "GetAllTablesOnPage.docx");
builder.CloseFile();
```

## Spreadsheet

Как использовать методы класса **ApiDocument** для таблиц и презентаций? Следовательно, не получились следующие примеры

Spreadsheet API/ApiDocument/AddElement
Spreadsheet API/ApiDocument/GetElement
Spreadsheet API/ApiDocument/GetElementsCount
Spreadsheet API/ApiDocument/Push
Spreadsheet API/ApiDocument/RemoveAllElements
Spreadsheet API/ApiDocument/RemoveElement

Presentation API/ApiDocument/AddElement
Presentation API/ApiDocument/GetElement
Presentation API/ApiDocument/GetElementsCount
Presentation API/ApiDocument/Push
Presentation API/ApiDocument/RemoveAllElements
Presentation API/ApiDocument/RemoveElement

>ApiDocument используется для CDE, и должно быть исключено из документации для CSE CPE

### Spreadsheet API/ApiRange/ForEach

`Вообще не поняла, для чего его используют.`

### Spreadsheet API/ApiRange/SetHidden

Не скрывает ячейки.

>Написал баг на Александра Трофимова

builder.CreateFile("xlsx");
var oWorksheet = Api.GetActiveSheet();
var oRange = oWorksheet.GetRange("A1:C1");
oRange.SetHidden(true);
oWorksheet.GetRange("A1").SetValue("1");
oWorksheet.GetRange("B1").SetValue("2");
oWorksheet.GetRange("C1").SetValue("3");
oWorksheet.GetRange("A3").SetValue("The values in cells A1:C1 are hidden.");
builder.SaveFile("xlsx", "SetHidden.xlsx");
builder.CloseFile();

 (Version: 5.6.3 (build:2)

### Spreadsheet API/ApiRange/GetHidden

`Возвращает тип null. Т.е. возникает ошибка.

>В develop возвращает в данном скрипте bool = false
>Перепроверить после фикса SetHidden()

builder.CreateFile("xlsx");
var oWorksheet = Api.GetActiveSheet();
var oRange = oWorksheet.GetRange("A1:C1");
oRange.SetHidden(true);
oWorksheet.GetRange("A1").SetValue("1");
oWorksheet.GetRange("B1").SetValue("2");
oWorksheet.GetRange("C1").SetValue("3");
oWorksheet.GetRange("A3").SetValue("The values in cells A1:C1 are hidden.");
var oHidden = oRange.GetHidden();
oWorksheet.GetRange("A4").SetValue("Hidden: " + oHidden);
builder.SaveFile("xlsx", "GetHidden.xlsx");
builder.CloseFile();

(Version: 5.6.3 (build:2)

### Spreadsheet API/ApiRange/SetOffset

`Не понимаю, что делает этот метод. Он вроде работает, но ничего не меняется.`

builder.CreateFile("xlsx");
var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("B3").SetValue("This is a sample text with cell offset specified.");
oWorksheet.GetRange("B3").SetOffset(2, 1);
builder.SaveFile("xlsx", "SetOffset.xlsx");
builder.CloseFile();

(Version: 5.6.3 (build:2)

### Spreadsheet API/ApiRange/SetRowHeight

Высота строки не меняется.

>Написал баг на Александра Трофимова

builder.CreateFile("xlsx");
var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("A1").SetRowHeight(32);
builder.SaveFile("xlsx", "SetRowHeight.xlsx");
builder.CloseFile();

(Version: 5.6.3 (build:2)

### **Spreadsheet** API/ApiWorksheet/ReplaceCurrentImage **Presentation** API/ApiPresentation/ReplaceCurrentImage

`Метод работает, но у меня получается сделать это только в два этапа. Т.е. сначала вставляем картинку, потом выходим из макроса, выделяем картинку и уже после этого используем метод ReplaceCurrentImage. Я не нашла метода, с помощью которого можно было бы выделить картинку.`

```js
builder.CreateFile("xlsx");
var oWorksheet = Api.GetActiveSheet();
oWorksheet.ReplaceCurrentImage("https://helpcenter.onlyoffice.com/images/Help/GettingStarted/Documents/big/EditDocument.png", 60 * 36000, 35 * 36000);
builder.SaveFile("xlsx", "ReplaceCurrentImage.xlsx");
builder.CloseFile();
```

(Version: 5.6.3 (build:2)

### Spreadsheet API/ApiChart/SetShowPointDataLabel

`Метод не работает.

```js
builder.CreateFile("xlsx");
var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 5, 2 * 36000, 1, 3 * 36000);
oChart.SetTitle('Financial Overview', 13);
oChart.SetShowPointDataLabel(1, 0, false, false, true, false);
builder.SaveFile("xlsx", "SetShowPointDataLabel.xlsx");
builder.CloseFile();
```

>Завел баг на Сергея Лузянина

Version: 6.0.0 (build:105)

### Spreadsheet API/ApiRange/Select

`Не работает.`

builder.CreateFile("xlsx");
var oWorksheet = Api.GetActiveSheet();
var oRange = oWorksheet.GetRange("A1:C1");
oRange.SetValue("1");
var oSelection = oRange.Select();
oSelection.SetFillColor(Api.CreateColorFromRGB(255, 224, 204));
builder.SaveFile("xlsx", "Select.xlsx");
builder.CloseFile();

>Символ вместо пробела «°»
>Результат у меня

А почему он не выполняет команду  oSelection.SetFillColor(Api.CreateColorFromRGB(255, 224, 204)) ? Что вообще возвращает этот метод? Видимо, не ApiRange?
SetFillColor должен возвращать либо true либо false (либо undefined я заводил баг недавно на решение вопроса о возвращении всеми методами bool значений)

Просто проблема в том, что такого примера, который ниже, будет недостаточно, потому что в demo этого выделения не видно.
Я так понял, идея примера в том что бы изменить цвет value, тем самым продемонстрировать работу метода, но что-то пошло не так
Посмотрю в пн внимательно

`Я вроде поняла причину. Метод  Select возвращает undefined. Следовательно, применить метод SetFillColor к тому, что возвращает Select, я не могу.`
`Подумаю, как еще можно показать работу метода.`

Version: 6.0.0 (build:105)

### Spreadsheet API/Api/Intersect

`Не работает.`

builder.CreateFile("xlsx");
var oWorksheet = Api.GetActiveSheet();
var oRange1 = oWorksheet.GetRange("A1:C5");
var oRange2 = oWorksheet.GetRange("B2:B4");
var oRange = Api.Intersect(oRange1, oRange2);
oRange.SetFillColor(Api.CreateColorFromRGB(255, 224, 204));
builder.SaveFile("xlsx", "GetCells.xlsx");
builder.CloseFile();

Version: 6.0.0 (build:105)

### Spreadsheet API/ApiComment/Delete

`Не работает.`

builder.CreateFile("xlsx");
var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("A1").SetValue("1");
var oRange = oWorksheet.GetRange("A1");
oRange.AddComment("This is just a number.");
var oComment = oRange.GetComment();
oComment.Delete();
oWorksheet.GetRange("A3").SetValue("The comment from the cell A1 was deleted.");
builder.SaveFile("xlsx", "Delete.xlsx");
builder.CloseFile();

Version: 6.0.0 (build:105)

### Spreadsheet API/ApiParagraph/Copy

`Не работает.`

```js
builder.CreateFile("xlsx");
var oWorksheet = Api.GetActiveSheet();
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(104, 155, 104));
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
var oDocContent = oShape.GetDocContent();
oDocContent.RemoveAllElements();
var oParagraph = Api.CreateParagraph();
oParagraph.AddText("This is just a sample text that was copied.");
oDocContent.Push(oParagraph);
var oCopyParagraph = oParagraph.Copy();
oDocContent.Push(oCopyParagraph);
builder.SaveFile("xlsx", "Copy.xlsx");
builder.CloseFile();
```

Version: 6.0.0 (build:105)

### Spreadsheet API/ApiWorksheet/SetActive

`Ошибка при вызове SetActive.`

builder.CreateFile("xlsx");
var oSheet = Api.AddSheet("New_sheet");
oSheet.SetActive();
oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("A1").SetValue("The current sheet is active.");
builder.SaveFile("xlsx", "SetActive.xlsx");
builder.CloseFile();

Version: 6.0.0 (build:105)

### Spreadsheet API/ApiTable

`Я не нашла метод для создания таблицы, поэтому методы для этого класса использовать не могу.`

### Spreadsheet API/ApiWorksheet

`При вызове методов GetPrintHeadings, SetPrintHeadings, GetPrintGridlines, SetPrintGridlines возникает ошибка.`

## Presentation

### Presentation API/Api/CreateGroup

Не понимаю, что делает этот метод.

>Метод для создание из массива объектов группы. По идее группы нужны для того чтобы обращаться с массивом объектов как с единым объектом.
>Класс есть, методы не реализованны. `Надобно исключить из документации до реализации методов.`

### Presentation API/ApiTable/Copy

`Ошибка при вызове метода Copy.`

builder.CreateFile("pptx");
var oPresentation = Api.GetPresentation();
var oTable = Api.CreateTable(2, 4);
oTable.SetPosition(608400, 1267200);
var oRow = oTable.GetRow(0);
var oCell = oRow.GetCell(0);
var oContent = oCell.GetContent();
var oParagraph = Api.CreateParagraph();
oParagraph.AddText("This is a table that was copied.");
oContent.Push(oParagraph);
var oSlide = oPresentation.GetSlideByIndex(0);
oSlide.RemoveAllObjects();
oSlide.AddObject(oTable);
var oCopyTable = oTable.Copy();
oSlide.AddObject(oCopyTable);
builder.SaveFile("pptx", "Copy.pptx");
builder.CloseFile();
Version: 6.0.0 (build:105)

### ApiDrawing

1. AddBreak

    ```js
    builder.CreateFile("docx");
    oDocument = Api.GetDocument();
    oParagraph = oDocument.GetElement(0);
    oFill = Api.CreateSolidFill(Api.CreateRGBColor(104, 155, 104));
    oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    oDrawing = Api.CreateShape("rect", 3212465, 963295, oFill, oStroke);
    oParagraph.AddDrawing(oDrawing);
    oDrawing.InsertParagraph("Added an LineBreak after this paragraph.", "before", false);
    oDrawing.AddBreak(1, "before");
    builder.SaveFile("docx", "AddBreak.docx");
    builder.CloseFile();
    ```

2. Delete

    ```js
    builder.CreateFile("docx");
    oDocument = Api.GetDocument();
    oParagraph = oDocument.GetElement(0);
    oFill = Api.CreateSolidFill(Api.CreateRGBColor(104, 155, 104));
    oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    oDrawing = Api.CreateShape("rect", 3212465, 963295, oFill, oStroke);
    oParagraph.AddDrawing(oDrawing);
    oDrawing.Delete();
    oParentParagraph.AddLineBreak();
    oParentParagraph.AddText("In this paragraph, the object Drawing has been deleted");
    builder.SaveFile("docx", "Delete.docx");
    builder.CloseFile();
    ```

3. Select

    ```js
    builder.CreateFile("docx");
    oDocument = Api.GetDocument();
    oParagraph = oDocument.GetElement(0);
    oFill = Api.CreateSolidFill(Api.CreateRGBColor(104, 155, 104));
    oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    oDrawing = Api.CreateShape("rect", 3212465, 963295, oFill, oStroke);
    oParagraph.AddDrawing(oDrawing);
    oDrawing.InsertParagraph("The Select property is applied to the drawing object", "before", false);
    oDrawing.Select();
    builder.SaveFile("docx", "Select.docx");
    builder.CloseFile();
    ```

4. ScaleHeight

    ```js
    builder.CreateFile("docx");
    oDocument = Api.GetDocument();
    oParagraph = oDocument.GetElement(0);
    for ( i = 3; i > 0; i-- ){
    oFill = Api.CreateSolidFill(Api.CreateRGBColor(104, 155, 104));
    oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    oDrawing = Api.CreateShape("cube", 3212465, 963295, oFill, oStroke);
    oParagraph.AddDrawing(oDrawing);
    oDrawing.ScaleHeight( i );} 
    builder.SaveFile("docx", "ScaleHeight.docx");
    builder.CloseFile();
    ```

5. ScaleWidth

    builder.CreateFile("docx");
    oDocument = Api.GetDocument();
    oParagraph = oDocument.GetElement(0);
        for (i = 1; i < 4; i++ ){
    oFill = Api.CreateSolidFill(Api.CreateRGBColor(104, 155, 104));
    oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    oDrawing = Api.CreateShape("cube", 963295, 963295, oFill, oStroke);
    oParagraph.AddDrawing(oDrawing);
    oDrawing.ScaleWidth( i );}
    builder.SaveFile("docx", "ScaleWidth.docx");
    builder.CloseFile();

6. GetParentParagraph

    builder.CreateFile("docx");
    oDocument = Api.GetDocument();
    oParagraph = oDocument.GetElement(0);
    oFill = Api.CreateSolidFill(Api.CreateRGBColor(104, 155, 104));
    oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    oDrawing = Api.CreateShape("rect", 3212465, 963295, oFill, oStroke);
    oParagraph.AddDrawing(oDrawing);
    oParentParagraph = oDrawing.GetParentParagraph();
    oParentParagraph.AddLineBreak();
    oParentParagraph.AddText("This is a parent paragraph");
    builder.SaveFile("docx", "GetParentParagraph.docx");
    builder.CloseFile();

7. GetParentTableCell

    builder.CreateFile("docx");
    oDocument = Api.GetDocument();
    oParagraph = Api.CreateParagraph();
    oTable = Api.CreateTable(3, 3);
    oTable.SetWidth("percent", 100);
    oFill = Api.CreateSolidFill(Api.CreateRGBColor(104, 155, 104));
    oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    oDrawing = Api.CreateShape("rect", 3212465, 963295, oFill, oStroke);
    oCell = oTable.GetCell(1, 1);
    oCell.GetContent().GetElement(0).AddDrawing(oDrawing);
    oDocument.Push(oTable);
    oParentCell = oDrawing.GetParentTableCell();
    oCellContent = oParentCell.GetContent().GetElement(0);
    oCellContent.AddLineBreak();
    oCellContent.AddText("This is a parent cell");
    builder.SaveFile("docx", "GetParentTableCell.docx");
    builder.CloseFile();

8. GetParentTable

    builder.CreateFile("docx");
    oDocument = Api.GetDocument();
    oParagraph = Api.CreateParagraph();
    oTable = Api.CreateTable(3, 3);
    oTable.SetWidth("percent", 100);
    oFill = Api.CreateSolidFill(Api.CreateRGBColor(104, 155, 104));
    oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    oDrawing = Api.CreateShape("rect", 3212465, 963295, oFill, oStroke);
    oCell = oTable.GetCell(1, 1);
    oCell.GetContent().GetElement(0).AddDrawing(oDrawing);
    oDocument.Push(oTable);
    oParentTable = oDrawing.GetParentTable();
    oTableStyle.SetBasedOn(oDocument.GetStyle("Table Grid"));
    oParentTable.SetStyle(oTableStyle);
    builder.SaveFile("docx", "GetParentTable.docx");
    builder.CloseFile();

9. InsertInContentControl

    builder.CreateFile("docx");
    oDocument = Api.GetDocument();
    oParagraph = oDocument.GetElement(0);
    oParagraph.AddText("This graphic object was wrapped in content control");
    oFill = Api.CreateSolidFill(Api.CreateRGBColor(104, 155, 104));
    oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    oDrawing = Api.CreateShape("rect", 3212465, 963295, oFill, oStroke);
    oParagraph.AddDrawing(oDrawing);
    oDrawing.InsertInContentControl(1);
    builder.SaveFile("docx", "InsertInContentControl.docx");
    builder.CloseFile();

10. InsertParagraph

    builder.CreateFile("docx");
    oDocument = Api.GetDocument();
    oParagraph = oDocument.GetElement(0);
    oFill = Api.CreateSolidFill(Api.CreateRGBColor(104, 155, 104));
    oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    oDrawing = Api.CreateShape("rect", 3212465, 963295, oFill, oStroke);
    oParagraph.AddDrawing(oDrawing);
    oDrawing.InsertParagraph("This is inserted paragraph.", "before", true);
    builder.SaveFile("docx", "InsertParagraph.docx");
    builder.CloseFile();

11. ScaleHeight и ScaleWidth

```js
    Методы, вроде, срабатывают. Потом я закрываю документ, открываю его снова, а там пусто.
    builder.CreateFile("docx");
    oDocument = Api.GetDocument();
    oParagraph = oDocument.GetElement(0);
    for ( i = 3; i > 0; i-- ){
    oFill = Api.CreateSolidFill(Api.CreateRGBColor(104, 155, 104));
    oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    oDrawing = Api.CreateShape("cube", 3212465, 963295, oFill, oStroke);
    oParagraph.AddDrawing(oDrawing);
    oDrawing.ScaleHeight( i );}
    builder.SaveFile("docx", "ScaleHeight.docx");
    builder.CloseFile();
```

```js
    builder.CreateFile("docx");
    oDocument = Api.GetDocument();
    oParagraph = oDocument.GetElement(0);
        for (i = 1; i < 4; i++ ){
    oFill = Api.CreateSolidFill(Api.CreateRGBColor(104, 155, 104));
    oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    oDrawing = Api.CreateShape("cube", 963295, 963295, oFill, oStroke);
    oParagraph.AddDrawing(oDrawing);
    oDrawing.ScaleWidth( i );}
    builder.SaveFile("docx", "ScaleWidth.docx");
    builder.CloseFile();
```

### ApiParagraph

1. Last

    ```js
    builder.CreateFile("docx");
    oDocument = Api.GetDocument();
    oParagraph = oDocument.GetElement(0);
    oRun_1 = Api.CreateRun();
    oRun_1.AddText("This is an Run with text. ");
    oParagraph.Push(oRun_1);
    oRun_2 = Api.CreateRun();
    oRun_2.AddText("And this is the last Run in the paragraph.");
    oParagraph.Push(oRun_2);
    oLastRun = oParagraph.Last();
    oLastRun.SetBold(true);
    builder.SaveFile("docx", "Last.docx");
    builder.CloseFile();
    ```

2. WrapInMailMergeField (для ApiParagraph и ApiRun). Я не особо поняла, что должно произойти при вызове этого метода, но в любом случае ничего не происходит.

>Я откатил изменения для этих методов на Api.onlyoffice
>Данные методы не заливали в прод тк функционал еще правят. Как он выйдет, объясню для чего данные методы.

```js
    (ApiRun)
    builder.CreateFile("docx");
    oDocument = Api.GetDocument();
    oParagraph = oDocument.GetElement(0);
    oRun = Api.CreateRun();
    oRun.AddText("Name");
    oParagraph.AddElement(oRun);
    oRun.WrapInMailMergeField();
    oParagraph.AddLineBreak();
    oRun = Api.CreateRun();
    oRun.AddText("Surname");
    oParagraph.AddElement(oRun);
    oRun.WrapInMailMergeField();
    builder.SaveFile(\"docx\", \"WrapInMailMergeField.docx\");
    builder.CloseFile();
```

```js
    (ApiParagraph)
    builder.CreateFile("docx");
    oDocument = Api.GetDocument();
    oParagraph = oDocument.GetElement(0);
    oParagraph.AddText("Paragraph wrapped in 'Mail Merge Field'");
    oParagraph.WrapInMailMergeField();
    builder.SaveFile("docx", "WrapInMailMergeField.docx");
    builder.CloseFile();
```

### ApiTableRow

1. Search

`Выделяет всю строку, а не конкретный объект в массиве.`

```js
    builder.CreateFile("docx");
    oDocument = Api.GetDocument();
    oTable = Api.CreateTable(3, 3);
    oRow = oTable.GetRow(0);
    oRow.GetCell(0).GetContent().GetElement(0).AddText("text");
    oRow.GetCell(1).GetContent().GetElement(0).AddText("text");
    oRow.GetCell(2).GetContent().GetElement(0).AddText("text");
    oDocument.Push(oTable);
    oRowSearch = oRow.Search("tex", true);
    oRowSearch[1].SetBold("true");
    builder.SaveFile("docx", "Search.docx");
    builder.CloseFile();
```
