# fluent-epplus

## Features
- Table Rendering
- Form Rendering
- Grouping columns header
- Handle inner collections as new table inside a 'cell'
- Few style options and capable of conditional styling

## To-Do's
- Tests
- Validate bad mapping
- Expose as extension of ExcelPackage
- More style options
- Create nuget package

## Examples

![alt text](https://github.com/mrxrsd/fluent-epplus/blob/main/docs/1.png?raw=true)

```csharp
var sheet1 = mapper.AddWorksheet().Label("Foo Sheet");
var mainTable = sheet1.AddTable<Foo>(t => t.Bind(Foos));

mainTable.MapProperty(x => x.Id).SetCaption("Id");

mainTable.MapGroup(m =>
{
    m.MapProperty(x => x.Description).SetCaption("Description").AutoFit();
    m.MapProperty(x => x.InnerFoo.Description).SetCaption("Inner Description").AutoFit();
    m.MapProperty(x => x.IsEnabled).SetHeaderStyle(s =>
    {
        s.BackgroundColor = Color.Black;
        s.FontColor = Color.Pink;
    }).SetStyle((style, dto) =>
    {
        style.FontColor = (dto.IsEnabled) ? Color.Green : Color.Red;
    }).SetCaption("Is Enabled");

}).SetCaption("General Info");
```

![alt text](https://github.com/mrxrsd/fluent-epplus/blob/main/docs/2.png?raw=true)

```csharp
var sheet2 = mapper.AddWorksheet().Label("Bar Sheet");
var mainTable = sheet2.AddTable<Bar>(t => t.Bind(Bars).AutoFit());
mainTable.MapProperty(x => x.Id).SetCaption("Id").AutoFit();
mainTable.MapProperty(x => x.Description).SetCaption("Description").AutoFit();

mainTable.MapProperty(x => x.CreatedAt, time => time.ToString("yyyy")).SetCaption("Year");
```

![alt text](https://github.com/mrxrsd/fluent-epplus/blob/main/docs/3.png?raw=true)
```csharp
var sheet = mapper.AddWorksheet().Label("FooBar Sheet").BlankAsNew();

var fooTable = sheet.AddTable<Foo>(t => t.Bind(Foos).StartAtRow(1));
fooTable.MapProperty(x => x.Id).SetCaption("Id");
fooTable.MapProperty(x => x.Description).SetCaption("Description");

var barTable = sheet.AddTable<Bar>(t => t.Bind(Bars).StartAtRow(5));
barTable.MapProperty(x => x.Id).SetCaption("Id");
barTable.MapProperty(x => x.Description).SetCaption("Description");
```

![alt text](https://github.com/mrxrsd/fluent-epplus/blob/main/docs/4.png?raw=true)
```csharp
var sheet = mapper.AddWorksheet().Label("Foo Form");

var form = sheet.AddForm<Foo>(f =>
{
    f.Bind(Foos)
     .StartAtRow(2)
     .SetCaption("Application")
     .AutoFit()
     .SetOrder(3);
});

form.MapProperty(m => m.Id).SetCaption("Id").AutoFit();
form.MapProperty(m => m.Description).SetCaption("Description").AutoFit();
```

![alt text](https://github.com/mrxrsd/fluent-epplus/blob/main/docs/5.png?raw=true)
```csharp
var sheet = mapper.AddWorksheet().Label("Complex Foo Sheet").AddTable<FooBars>(t => t.Bind(FoosBars));

sheet.MapProperty(x => x.Id).SetCaption("Id").AutoFit();
sheet.MapProperty(x => x.Description).SetCaption("Description").AutoFit();
sheet.MapProperty(x => x.CreatedAt, time => time.ToString("yyyy-MM-dd")).SetCaption("Create At");

sheet.MapTable(y => y.Foos, x =>
    {
        x.MapProperty(p => p.Id).SetCaption("Id").AutoFit();
        x.MapProperty(p => p.Description).SetCaption("Description").AutoFit();
        x.MapProperty(p => p.IsEnabled).SetCaption("Is Enabled").AutoFit();

    }).ShowHeaderPerRow()
    .SetCaption("Inner Table");
```
