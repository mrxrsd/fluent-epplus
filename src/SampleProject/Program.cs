using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using FluentEpplus;

namespace SampleProject
{

    public class InnerFoo
    {
        public string Description { get; set; }
    }

    public class Foo
    {
        public int Id { get; set; }
        public string Description { get; set; }
        public bool IsEnabled { get; set; }
        public InnerFoo InnerFoo { get; set; }
        public List<Bar> Bars { get; set; }
    }

    public class Bar
    {
        public int Id { get; set; }
        public string Description { get; set; }
        public DateTime CreatedAt { get; set; }
    }

    public class FooBars
    {
        public int Id { get; set; }
        public string Description { get; set; }
        public DateTime CreatedAt { get; set; }
        public List<Foo> Foos { get; set; }
    }
    

    class Program
    {
        public const string OUTPUT_PATH = @"C:\temp\";
        public static List<Foo> Foos;
        public static List<Bar> Bars;
        public static List<FooBars> FoosBars;

        static void Main(string[] args)
        {
            Bars = new List<Bar>
            {
                new Bar { Id = 1, Description = "Bar One", CreatedAt = DateTime.Now},
                new Bar { Id = 2, Description = "Bar Two", CreatedAt = DateTime.Now},
            };

            Foos = new List<Foo>
            {
                new Foo { Id=1, Description = "Foo One loooooooooooooooongggg", InnerFoo = new InnerFoo { Description = "Inner Description"}, IsEnabled = true, Bars = Bars},
                new Foo { Id=2, Description = "Foo Two", IsEnabled = false, Bars = Bars},
                new Foo { Id=3, Description = "Foo Three", IsEnabled = true, Bars = Bars},
                new Foo { Id=4, Description = "Foo Four", IsEnabled = true, Bars = Bars},

            };

            FoosBars = new List<FooBars>
            {
                new FooBars { Id = 1, Description = "FooBar One", CreatedAt = DateTime.Now, Foos = Foos},
                new FooBars { Id = 2, Description = "FooBar Two", CreatedAt = DateTime.Now, Foos = Foos},
            };

            var mapper = new ExcelMap();

            Sample1(mapper);
            Sample2(mapper);
            Sample3(mapper);
            Sample4(mapper);
            Sample5(mapper);

            var filename = $"extractFile{DateTime.Now.Ticks}.xlsx";
            Extract(Path.Combine(OUTPUT_PATH,filename), mapper.GetExcelPackage().GetAsByteArray());
        }

        private static void Sample5(ExcelMap mapper)
        {
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
        }

        private static void Sample4(ExcelMap mapper)
        {
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
        }

        private static void Sample3(ExcelMap mapper)
        {
            var sheet = mapper.AddWorksheet().Label("FooBar Sheet").BlankAsNew();

            var fooTable = sheet.AddTable<Foo>(t => t.Bind(Foos).StartAtRow(1));
            fooTable.MapProperty(x => x.Id).SetCaption("Id");
            fooTable.MapProperty(x => x.Description).SetCaption("Description");

            var barTable = sheet.AddTable<Bar>(t => t.Bind(Bars).StartAtRow(5));
            barTable.MapProperty(x => x.Id).SetCaption("Id");
            barTable.MapProperty(x => x.Description).SetCaption("Description");
        }

        private static void Sample1(ExcelMap mapper)
        {
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
        }

        private static void Sample2(ExcelMap mapper)
        {
            var sheet2 = mapper.AddWorksheet().Label("Bar Sheet");
            var mainTable = sheet2.AddTable<Bar>(t => t.Bind(Bars).AutoFit());
            mainTable.MapProperty(x => x.Id).SetCaption("Id").AutoFit();
            mainTable.MapProperty(x => x.Description).SetCaption("Description").AutoFit();

            mainTable.MapProperty(x => x.CreatedAt, time => time.ToString("yyyy")).SetCaption("Year");
        }


        private static void Extract(string filename, byte[] stream)
        {
            BinaryWriter bw = new BinaryWriter(File.Open(filename, FileMode.CreateNew));
            bw.Write(stream);
            bw.Flush();
            bw.Close();

        }
    }

}
