using System;
using System.Collections.Generic;
using System.Drawing;
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
            //Sample2(mapper);

            var filename = "extractFile" + DateTime.Now.Ticks;
            mapper.Extract(OUTPUT_PATH, filename, true);
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
    }

}
