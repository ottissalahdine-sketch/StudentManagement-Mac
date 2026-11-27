using System.Globalization;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;

namespace StudentManagement.Desktop;

public sealed class StudentFormWindow : Window
{
    private readonly string? _originalNo;

    private readonly TextBox no=MainWindow.Field("Student No");
    private readonly TextBox name=MainWindow.Field("Name");
    private readonly ComboBox gender=new(){MinHeight=40,HorizontalAlignment=HorizontalAlignment.Stretch};
    private readonly TextBox iquama=MainWindow.Field("Iquama");
    private readonly DatePicker dob=new(){MinHeight=40,HorizontalAlignment=HorizontalAlignment.Stretch};
    private readonly ComboBox grade=new(){MinHeight=40,HorizontalAlignment=HorizontalAlignment.Stretch};
    private readonly ComboBox room=new(){MinHeight=40,HorizontalAlignment=HorizontalAlignment.Stretch};
    private readonly TextBox p1=MainWindow.Field("Phone No");
    private readonly TextBox p2=MainWindow.Field("Phone No2");
    private readonly TextBox p3=MainWindow.Field("Phone No3");
    private readonly TextBox p4=MainWindow.Field("Phone No4");

    public bool Saved {get;private set;}

    private static readonly string[] Grades={
        "NURS","KG1","KG2","KG3",
        "GR01","GR02","GR03","GR04","GR05","GR06",
        "GR07","GR08","GR09","GR10","GR11","GR12"
    };

    private static readonly string[] Rooms={"A","B","C","D","E","X"};

    public StudentFormWindow(StudentRecord? s=null)
    {
        _originalNo=s?.StudentNo;
        Title=s is null?"Add Student":"Edit Student";
        Width=820; Height=690;
        MinWidth=720; MinHeight=620;
        WindowStartupLocation=WindowStartupLocation.CenterOwner;
        Background=MainWindow.Brush("#EEF3F8");

        gender.ItemsSource=new[]{"Male","Female"};
        grade.ItemsSource=Grades;
        room.ItemsSource=Rooms;

        if(s is not null)
        {
            no.Text=s.StudentNo;
            name.Text=s.Name;
            gender.SelectedItem=NormalizeGender(s.Gender);
            iquama.Text=s.Iquama;
            dob.SelectedDate=ParseDate(s.DateOfBirth);
            grade.SelectedItem=SelectOrKeep(Grades,s.Grade);
            room.SelectedItem=SelectOrKeep(Rooms,s.Room);
            p1.Text=s.Phone1;
            p2.Text=s.Phone2;
            p3.Text=s.Phone3;
            p4.Text=s.Phone4;
        }
        else
        {
            gender.SelectedIndex=0;
        }

        Content=Build();
    }

    private Control Build()
    {
        var save=MainWindow.PrimaryButton("SAVE");
        save.Click+=async(_,_)=>await Save();

        var clear=MainWindow.LightButton("CLEAR");
        clear.Click+=(_,_)=>ClearForm();

        var back=MainWindow.SecondaryButton("← BACK");
        back.Click+=(_,_)=>Close();

        var header=new Grid{ColumnDefinitions=ColumnDefinitions.Parse("*,Auto")};
        header.Children.Add(new StackPanel{Spacing=2,Children={
            MainWindow.Header(Title.ToUpperInvariant(),26),
            new TextBlock{
                Text=_originalNo is null?"Create a new student record":"Update the selected student record",
                Foreground=MainWindow.Brush("#64748B"),FontSize=12
            }
        }});
        var buttons=MainWindow.Row(save,clear,back);
        buttons.SetValue(Grid.ColumnProperty,1);
        header.Children.Add(buttons);

        var form=new Grid{
            ColumnDefinitions=ColumnDefinitions.Parse("*,*"),
            RowDefinitions=RowDefinitions.Parse("Auto,Auto,Auto,Auto,Auto,Auto"),
            ColumnSpacing=18,
            RowSpacing=12
        };

        AddField(form,0,0,"STUDENT NO",no);
        AddField(form,0,1,"NAME",name);
        AddField(form,1,0,"GENDER",gender);
        AddField(form,1,1,"IQUAMA",iquama);
        AddField(form,2,0,"DATE OF BIRTH",dob);
        AddField(form,2,1,"GRADE",grade);
        AddField(form,3,0,"ROOM",room);
        AddField(form,3,1,"PHONE NO",p1);
        AddField(form,4,0,"PHONE NO 2",p2);
        AddField(form,4,1,"PHONE NO 3",p3);
        AddField(form,5,0,"PHONE NO 4",p4);

        var note=new Border{
            Background=MainWindow.Brush("#F0F7F6"),
            BorderBrush=MainWindow.Brush("#B7D9D4"),
            BorderThickness=new Thickness(1),
            CornerRadius=new CornerRadius(12),
            Padding=new Thickness(14,10),
            Child=new TextBlock{
                Text="Date of Birth is saved as dd/MM/yyyy. Grade and Room use the controlled school lists.",
                Foreground=MainWindow.Brush("#355E59"),FontSize=11,TextWrapping=TextWrapping.Wrap
            }
        };
        note.SetValue(Grid.RowProperty,5);
        note.SetValue(Grid.ColumnProperty,1);
        form.Children.Add(note);

        var card=MainWindow.Card(form,new Thickness(22),"#FFFFFF");

        return new Grid{
            Margin=new Thickness(26,22),
            RowDefinitions=RowDefinitions.Parse("Auto,*"),
            Children={
                header,
                card
            }
        }.WithChildRow(card,1,new Thickness(0,18,0,0));
    }

    private static void AddField(Grid grid,int row,int col,string label,Control control)
    {
        var panel=new StackPanel{Spacing=6,Children={
            new TextBlock{Text=label,FontSize=11,FontWeight=Avalonia.Media.FontWeight.Bold,Foreground=MainWindow.Brush("#475569")},
            control
        }};
        panel.SetValue(Grid.RowProperty,row);
        panel.SetValue(Grid.ColumnProperty,col);
        grid.Children.Add(panel);
    }

    private async Task Save()
    {
        try
        {
            var genderValue=gender.SelectedItem?.ToString()??"";
            var gradeValue=grade.SelectedItem?.ToString()??"";
            var roomValue=room.SelectedItem?.ToString()??"";
            var dobValue=dob.SelectedDate?.ToString("dd/MM/yyyy",CultureInfo.InvariantCulture)??"";

            if(string.IsNullOrWhiteSpace(_originalNo))
                ExcelRepository.AddStudent(no.Text??"",name.Text??"",genderValue,iquama.Text??"",dobValue,gradeValue,roomValue,p1.Text??"",p2.Text??"",p3.Text??"",p4.Text??"");
            else
                ExcelRepository.UpdateStudent(_originalNo,no.Text??"",name.Text??"",genderValue,iquama.Text??"",dobValue,gradeValue,roomValue,p1.Text??"",p2.Text??"",p3.Text??"",p4.Text??"");

            Saved=true;
            Close();
        }
        catch(Exception ex)
        {
            await Ui.Alert(this,"Unable to Save Student",ex.Message);
        }
    }

    private void ClearForm()
    {
        no.Text="";
        name.Text="";
        gender.SelectedIndex=0;
        iquama.Text="";
        dob.SelectedDate=null;
        grade.SelectedIndex=-1;
        room.SelectedIndex=-1;
        p1.Text="";p2.Text="";p3.Text="";p4.Text="";
        no.Focus();
    }

    private static DateTimeOffset? ParseDate(string? value)
    {
        if(string.IsNullOrWhiteSpace(value)) return null;
        string[] formats={"dd/MM/yyyy","d/M/yyyy","dd-MM-yyyy","d-M-yyyy","yyyy-MM-dd"};
        if(DateTime.TryParseExact(value.Trim(),formats,CultureInfo.InvariantCulture,DateTimeStyles.None,out var d))
            return new DateTimeOffset(d);
        if(DateTime.TryParse(value,CultureInfo.InvariantCulture,DateTimeStyles.None,out d))
            return new DateTimeOffset(d);
        return null;
    }

    private static string? SelectOrKeep(IEnumerable<string> items,string? value)
    {
        var v=(value??"").Trim();
        return items.FirstOrDefault(x=>x.Equals(v,StringComparison.OrdinalIgnoreCase));
    }

    private static string NormalizeGender(string? value)
    {
        var v=(value??"").Trim();
        if(v.Equals("F",StringComparison.OrdinalIgnoreCase)||v.Equals("Female",StringComparison.OrdinalIgnoreCase)||v.Equals("Girl",StringComparison.OrdinalIgnoreCase))
            return "Female";
        return "Male";
    }
}

internal static class StudentFormGridExtensions
{
    public static Grid WithChildRow(this Grid grid,Control child,int row,Thickness margin)
    {
        child.SetValue(Grid.RowProperty,row);
        child.Margin=margin;
        return grid;
    }
}
