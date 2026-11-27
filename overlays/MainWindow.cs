using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Data;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using ClosedXML.Excel;

namespace StudentManagement.Desktop;

public sealed class MainWindow : Window
{
    private readonly StackPanel _menu = new(){Spacing=12,Width=560,HorizontalAlignment=HorizontalAlignment.Center};

    public MainWindow()
    {
        Title="Student Management";
        Width=920; Height=780; MinWidth=760; MinHeight=650;
        WindowStartupLocation=WindowStartupLocation.CenterScreen;
        Background=Brush("#EEF2F7");
        Content=BuildHome();
    }

    private Control BuildHome()
    {
        var logo=new TextBlock{Text="STUDENT MANAGEMENT",FontSize=30,FontWeight=FontWeight.Bold,HorizontalAlignment=HorizontalAlignment.Center,Foreground=Brush("#0E2554")};
        var school=new TextBlock{Text=AppSettingsManager.Current.SchoolName,FontSize=15,HorizontalAlignment=HorizontalAlignment.Center,Foreground=Brush("#475569")};
        AddHome("ADD STUDENT", async()=>{var w=new StudentFormWindow();await w.ShowDialog(this);});
        AddHome("SEARCH STUDENT", async()=>await new StudentsWindow().ShowDialog(this));
        AddHome("DASHBOARD", async()=>await new DashboardWindow().ShowDialog(this));
        AddHome("ACTIVITY LOG", async()=>await new ActivityLogWindow().ShowDialog(this));
        AddHome("DATA QUALITY", async()=>await new DataQualityWindow().ShowDialog(this));
        AddHome("IMPORT STUDENTS", async()=>await new ImportWizardWindow().ShowDialog(this));
        AddHome("SETTINGS", async()=>await new SettingsWindow().ShowDialog(this));
        AddHome("DELETED STUDENTS", async()=>await new DeletedStudentsWindow().ShowDialog(this));
        AddHome("EXIT", async()=>{if(await Ui.Confirm(this,"Exit","Exit Student Management?")) Close();});
        var card=Card(_menu,new Thickness(44),"#FFFFFF");
        return new ScrollViewer{Content=new StackPanel{Margin=new Thickness(40,26),Spacing=16,Children={logo,school,card}}};
    }

    private void AddHome(string text,Func<Task> action)
    {
        var b=PrimaryButton(text); b.Height=58; b.HorizontalAlignment=HorizontalAlignment.Stretch; b.FontSize=14;
        b.Click+=async(_,_)=>await action(); _menu.Children.Add(b);
    }

    internal static SolidColorBrush Brush(string hex)=>new(Color.Parse(hex));
    internal static Border Card(Control c,Thickness padding,string bg="#F8FAFC")=>new(){Child=c,Padding=padding,CornerRadius=new CornerRadius(16),Background=Brush(bg),BorderBrush=Brush("#CBD5E1"),BorderThickness=new Thickness(1)};
    internal static Button PrimaryButton(string text)=>new(){Content=text,Padding=new Thickness(18,10),Background=Brush("#00796B"),Foreground=Brush("#FFFFFF"),BorderThickness=new Thickness(0),CornerRadius=new CornerRadius(10),MinWidth=120};
    internal static Button SecondaryButton(string text)=>new(){Content=text,Padding=new Thickness(16,9),Background=Brush("#0E2554"),Foreground=Brush("#FFFFFF"),BorderThickness=new Thickness(0),CornerRadius=new CornerRadius(10),MinWidth=105};
    internal static Button LightButton(string text)=>new(){Content=text,Padding=new Thickness(16,9),Background=Brush("#F8FAFC"),Foreground=Brush("#0F172A"),BorderBrush=Brush("#CBD5E1"),BorderThickness=new Thickness(1),CornerRadius=new CornerRadius(10),MinWidth=105};
    internal static TextBlock Header(string t,double s=24)=>new(){Text=t,FontSize=s,FontWeight=FontWeight.Bold,Foreground=Brush("#0F172A")};
    internal static TextBox Field(string hint="")=>new(){Watermark=hint,MinHeight=38,Padding=new Thickness(10,7)};
    internal static StackPanel Row(params Control[] cs){var p=new StackPanel{Orientation=Orientation.Horizontal,Spacing=8};foreach(var c in cs)p.Children.Add(c);return p;}
    internal static string ExportPath(string stem,string ext)
    {
        var f=AppSettingsManager.Current.DefaultExportFolder;
        if(string.IsNullOrWhiteSpace(f)) f=Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        Directory.CreateDirectory(f);
        var safe=string.Concat(stem.Select(c=>Path.GetInvalidFileNameChars().Contains(c)?'_':c));
        return Path.Combine(f,$"{safe}_{DateTime.Now:yyyyMMdd_HHmmss}.{ext}");
    }
    internal static void AddStudentColumns(DataGrid g)
    {
        Add(g,"Student No","StudentNo",105); Add(g,"Name","Name",220); Add(g,"Gender","Gender",90); Add(g,"Iquama","Iquama",130);
        Add(g,"Date of Birth","DateOfBirth",120); Add(g,"Grade","Grade",95); Add(g,"Room","Room",75); Add(g,"Phone No","Phone1",130);
        Add(g,"Phone No2","Phone2",130); Add(g,"Phone No3","Phone3",130); Add(g,"Phone No4","Phone4",130);
    }
    internal static void Add(DataGrid g,string h,string p,double w)=>g.Columns.Add(new DataGridTextColumn{Header=h,Binding=new Binding(p),Width=new DataGridLength(w)});
    internal static DataGrid GridBase()=>new(){AutoGenerateColumns=false,IsReadOnly=true,GridLinesVisibility=DataGridGridLinesVisibility.All,SelectionMode=DataGridSelectionMode.Single,HeadersVisibility=DataGridHeadersVisibility.All};
    internal static async Task OpenPdfAfterExport(Window owner,string path)
    {
        try{Process.Start(new ProcessStartInfo("open",$"\"{path}\""){UseShellExecute=false});}
        catch{await Ui.Alert(owner,"PDF Created",path);}
    }
    private void Logout()
    {
        var login=new LoginWindow(); if(Application.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime d)d.MainWindow=login; login.Show(); Close();
    }
}

public sealed class DashboardWindow : Window
{
    private DataTable _all=new();
    private string _gender="All",_grade="All",_room="All";
    private readonly ComboBox gender=new(){MinWidth=150},grade=new(){MinWidth=150},room=new(){MinWidth=150};
    private readonly WrapPanel kpis=new();
    private readonly StackPanel genderChart=new(){Spacing=8},gradeChart=new(){Spacing=6},roomChart=new(){Spacing=6};
    private readonly DataGrid summary=MainWindow.GridBase();

    public DashboardWindow()
    {
        Title="Dashboard"; Width=1320; Height=860; MinWidth=1020; MinHeight=700; WindowStartupLocation=WindowStartupLocation.CenterOwner; Background=MainWindow.Brush("#E6EDF5");
        MainWindow.Add(summary,"Grade","Grade",120); MainWindow.Add(summary,"Male","Male",90); MainWindow.Add(summary,"Female","Female",90); MainWindow.Add(summary,"Total","Total",90);
        Content=Build(); Load();
    }
    private Control Build()
    {
        var refresh=MainWindow.PrimaryButton("REFRESH"); refresh.Click+=(_,_)=>Load();
        var clear=MainWindow.LightButton("CLEAR FILTERS"); clear.Click+=(_,_)=>{gender.SelectedIndex=0;grade.SelectedIndex=0;room.SelectedIndex=0;};
        var close=MainWindow.SecondaryButton("CLOSE"); close.Click+=(_,_)=>Close();
        gender.SelectionChanged+=(_,_)=>{_gender=gender.SelectedItem?.ToString()??"All";Render();};
        grade.SelectionChanged+=(_,_)=>{_grade=grade.SelectedItem?.ToString()??"All";Render();};
        room.SelectionChanged+=(_,_)=>{_room=room.SelectedItem?.ToString()??"All";Render();};
        var filters=MainWindow.Card(new StackPanel{Spacing=8,Children={new TextBlock{Text="FILTERS",FontWeight=FontWeight.Bold},MainWindow.Row(new TextBlock{Text="Gender",VerticalAlignment=VerticalAlignment.Center},gender,new TextBlock{Text="Grade",VerticalAlignment=VerticalAlignment.Center},grade,new TextBlock{Text="Room",VerticalAlignment=VerticalAlignment.Center},room,clear)}},new Thickness(16));
        var charts=new Grid{ColumnDefinitions=ColumnDefinitions.Parse("*,*,*")};
        var g1=ChartCard("GENDER DISTRIBUTION",genderChart); var g2=ChartCard("GRADE DISTRIBUTION",gradeChart); var g3=ChartCard("ROOM DISTRIBUTION",roomChart);
        g2.SetValue(Grid.ColumnProperty,1);g3.SetValue(Grid.ColumnProperty,2);charts.Children.Add(g1);charts.Children.Add(g2);charts.Children.Add(g3);
        var root=new StackPanel{Margin=new Thickness(24),Spacing=14,Children={MainWindow.Header("DASHBOARD",28),filters,kpis,charts,new TextBlock{Text="GRADE SUMMARY",FontWeight=FontWeight.Bold,FontSize=16},new Border{Height=260,Child=summary},MainWindow.Row(refresh,close)}};
        return new ScrollViewer{Content=root};
    }
    private Border ChartCard(string title,StackPanel body)=>MainWindow.Card(new StackPanel{Spacing=10,Children={new TextBlock{Text=title,FontWeight=FontWeight.Bold},body}},new Thickness(14),"#FFFFFF");
    private void Load()
    {
        try{_all=ExcelRepository.LoadStudents();}
        catch(Exception ex){_all=new DataTable(); _=Ui.Alert(this,"Database Error",ex.Message);return;}
        var grades=new[]{"All"}.Concat(Values("Grade")).ToList(); var rooms=new[]{"All"}.Concat(Values("Room")).ToList();
        gender.ItemsSource=new[]{"All","Male","Female"}; grade.ItemsSource=grades; room.ItemsSource=rooms;
        gender.SelectedItem=_gender;grade.SelectedItem=_grade;room.SelectedItem=_room; if(gender.SelectedItem is null)gender.SelectedIndex=0;if(grade.SelectedItem is null)grade.SelectedIndex=0;if(room.SelectedItem is null)room.SelectedIndex=0;
        Render();
    }
    private IEnumerable<string> Values(string c)=>_all.Columns.Contains(c)?_all.Rows.Cast<DataRow>().Select(r=>Convert.ToString(r[c])?.Trim()??"").Where(x=>x.Length>0).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x=>x):Enumerable.Empty<string>();
    private IEnumerable<DataRow> Filtered()=>_all.Rows.Cast<DataRow>().Where(r=>MatchGender(r)&&Match(r,"Grade",_grade)&&Match(r,"Room",_room));
    private bool Match(DataRow r,string c,string v)=>v=="All"||string.Equals(Convert.ToString(r[c])?.Trim(),v,StringComparison.OrdinalIgnoreCase);
    private bool MatchGender(DataRow r){if(_gender=="All")return true;var v=Convert.ToString(r["Gender"])?.Trim()??"";return _gender=="Male"?(v.Equals("M",StringComparison.OrdinalIgnoreCase)||v.Equals("Male",StringComparison.OrdinalIgnoreCase)):(v.Equals("F",StringComparison.OrdinalIgnoreCase)||v.Equals("Female",StringComparison.OrdinalIgnoreCase));}
    private void Render()
    {
        if(_all.Columns.Count==0)return; var rows=Filtered().ToList(); int total=rows.Count,male=rows.Count(IsMale),female=rows.Count(IsFemale);
        kpis.Children.Clear(); foreach(var kv in new[]{("TOTAL STUDENTS",total.ToString("#,0")),("MALE",male.ToString("#,0")),("FEMALE",female.ToString("#,0")),("GRADES",rows.Select(r=>S(r,"Grade")).Where(x=>x.Length>0).Distinct().Count().ToString()),("ROOMS",rows.Select(r=>S(r,"Room")).Where(x=>x.Length>0).Distinct().Count().ToString())})kpis.Children.Add(Kpi(kv.Item1,kv.Item2));
        FillChart(genderChart,new[]{("Male",male),("Female",female)},Math.Max(1,total));
        var gc=rows.GroupBy(r=>S(r,"Grade")).Where(g=>g.Key.Length>0).OrderBy(g=>g.Key).Select(g=>(g.Key,g.Count())).ToArray();FillChart(gradeChart,gc,Math.Max(1,gc.Select(x=>x.Item2).DefaultIfEmpty(1).Max()));
        var rc=rows.GroupBy(r=>S(r,"Room")).Where(g=>g.Key.Length>0).OrderBy(g=>g.Key).Select(g=>(g.Key,g.Count())).ToArray();FillChart(roomChart,rc,Math.Max(1,rc.Select(x=>x.Item2).DefaultIfEmpty(1).Max()));
        summary.ItemsSource=rows.GroupBy(r=>S(r,"Grade")).Where(g=>g.Key.Length>0).OrderBy(g=>g.Key).Select(g=>new GradeSummary{Grade=g.Key,Male=g.Count(IsMale),Female=g.Count(IsFemale),Total=g.Count()}).ToList();
    }
    private static bool IsMale(DataRow r){var v=S(r,"Gender");return v.Equals("M",StringComparison.OrdinalIgnoreCase)||v.Equals("Male",StringComparison.OrdinalIgnoreCase);}
    private static bool IsFemale(DataRow r){var v=S(r,"Gender");return v.Equals("F",StringComparison.OrdinalIgnoreCase)||v.Equals("Female",StringComparison.OrdinalIgnoreCase);}
    private static string S(DataRow r,string c)=>r.Table.Columns.Contains(c)?Convert.ToString(r[c])?.Trim()??"":"";
    private Border Kpi(string l,string v)=>MainWindow.Card(new StackPanel{Spacing=4,Children={new TextBlock{Text=l,Foreground=MainWindow.Brush("#64748B")},new TextBlock{Text=v,FontSize=28,FontWeight=FontWeight.Bold}}},new Thickness(16),"#FFFFFF") with {Width=225,Margin=new Thickness(0,0,10,0)};
    private void FillChart(StackPanel p,IEnumerable<(string,int)> data,int max){p.Children.Clear();foreach(var (label,value) in data.Take(16)){var bar=new ProgressBar{Minimum=0,Maximum=max,Value=value,Height=12};p.Children.Add(new Grid{ColumnDefinitions=ColumnDefinitions.Parse("90,*,45"),Children={new TextBlock{Text=label,TextTrimming=TextTrimming.CharacterEllipsis},bar,new TextBlock{Text=value.ToString(),HorizontalAlignment=HorizontalAlignment.Right}}});bar.SetValue(Grid.ColumnProperty,1);((Grid)p.Children[^1]).Children[^1].SetValue(Grid.ColumnProperty,2);}}
    private sealed class GradeSummary{public string Grade{get;set;}="";public int Male{get;set;}public int Female{get;set;}public int Total{get;set;}}
}

public sealed class StudentsWindow : Window
{
    private readonly ObservableCollection<StudentRecord> shown=new(); private List<StudentRecord> all=new(); private readonly DataGrid grid=MainWindow.GridBase(); private readonly TextBox search=MainWindow.Field("Search student no, name, grade, room, iqama, phone...");
    public StudentsWindow(){Title="Search Student";Width=1380;Height=820;MinWidth=980;MinHeight=620;WindowStartupLocation=WindowStartupLocation.CenterOwner;MainWindow.AddStudentColumns(grid);grid.ItemsSource=shown;Content=Build();Load();}
    private Control Build(){search.PropertyChanged+=(_,e)=>{if(e.Property==TextBox.TextProperty)Apply();};var add=MainWindow.PrimaryButton("ADD STUDENT");add.Click+=async(_,_)=>{var w=new StudentFormWindow();await w.ShowDialog(this);if(w.Saved)Load();};var edit=MainWindow.LightButton("EDIT");edit.Click+=async(_,_)=>await Edit();var profile=MainWindow.LightButton("PROFILE");profile.Click+=async(_,_)=>await Profile();var del=MainWindow.LightButton("DELETE");del.Click+=async(_,_)=>await Delete();var xls=MainWindow.LightButton("EXPORT EXCEL");xls.Click+=async(_,_)=>await Export("xlsx");var pdf=MainWindow.LightButton("EXPORT PDF");pdf.Click+=async(_,_)=>await Export("pdf");var close=MainWindow.SecondaryButton("CLOSE");close.Click+=(_,_)=>Close();var root=new Grid{Margin=new Thickness(22),RowDefinitions=RowDefinitions.Parse("Auto,Auto,*")};var head=new StackPanel{Spacing=10,Children={MainWindow.Header("SEARCH STUDENT"),search,MainWindow.Row(add,edit,profile,del,xls,pdf,close)}};root.Children.Add(head);var host=new Border{Margin=new Thickness(0,14,0,0),Child=grid};host.SetValue(Grid.RowProperty,2);root.Children.Add(host);return root;}
    private void Load(){try{all=ExcelRepository.LoadStudents().Rows.Cast<DataRow>().Select(StudentRecord.From).ToList();Apply();}catch(Exception ex){_=Ui.Alert(this,"Database Error",ex.Message);}}
    private void Apply(){var q=(search.Text??"").Trim();shown.Clear();foreach(var s in all.Where(s=>q.Length==0||$"{s.StudentNo} {s.Name} {s.Gender} {s.Iquama} {s.DateOfBirth} {s.Grade} {s.Room} {s.Phone1} {s.Phone2} {s.Phone3} {s.Phone4}".Contains(q,StringComparison.OrdinalIgnoreCase)))shown.Add(s);}
    private async Task Edit(){if(grid.SelectedItem is not StudentRecord s){await Ui.Alert(this,"Edit","Select a student first.");return;}var w=new StudentFormWindow(s);await w.ShowDialog(this);if(w.Saved)Load();}
    private async Task Profile(){if(grid.SelectedItem is not StudentRecord s){await Ui.Alert(this,"Profile","Select a student first.");return;}await new StudentProfileWindow(s).ShowDialog(this);Load();}
    private async Task Delete(){if(grid.SelectedItem is not StudentRecord s){await Ui.Alert(this,"Delete","Select a student first.");return;}if(await Ui.Confirm(this,"Delete Student",$"Move {s.StudentNo} - {s.Name} to Deleted Students?")){try{ExcelRepository.DeleteStudent(s.StudentNo);Load();}catch(Exception ex){await Ui.Alert(this,"Delete Error",ex.Message);}}}
    private async Task Export(string ext){try{var table=ExcelRepository.LoadStudents();var path=MainWindow.ExportPath("Student_Search",ext);if(ext=="xlsx")ReportExportService.ExportStudentsExcel(path,table,"Student Search",$"Records: {table.Rows.Count:#,0}");else ReportExportService.ExportStudentsPdf(path,table,"Student Search",$"Records: {table.Rows.Count:#,0}");await Ui.Alert(this,"Export Complete",path);}catch(Exception ex){await Ui.Alert(this,"Export Error",ex.Message);}}
}

public sealed class StudentProfileWindow : Window
{
    private readonly StudentRecord student; private readonly DataGrid activity=MainWindow.GridBase();
    public StudentProfileWindow(StudentRecord s){student=s;Title="Student Profile";Width=1180;Height=820;MinWidth=900;MinHeight=620;WindowStartupLocation=WindowStartupLocation.CenterOwner;MainWindow.Add(activity,"Date / Time","DateTimeText",170);MainWindow.Add(activity,"Action","Action",140);MainWindow.Add(activity,"Details","Details",520);Content=Build();LoadActivity();}
    private Control Build(){var details=new Grid{ColumnDefinitions=ColumnDefinitions.Parse("160,*,160,*"),RowDefinitions=RowDefinitions.Parse("Auto,Auto,Auto,Auto,Auto")};AddD(details,0,0,"Student No",student.StudentNo);AddD(details,0,1,"Name",student.Name);AddD(details,1,0,"Gender",student.Gender);AddD(details,1,1,"Iquama",student.Iquama);AddD(details,2,0,"Date of Birth",student.DateOfBirth);AddD(details,2,1,"Grade",student.Grade);AddD(details,3,0,"Room",student.Room);AddD(details,3,1,"Phone",student.Phone1);AddD(details,4,0,"Phone 2",student.Phone2);AddD(details,4,1,"Phone 3 / 4",$"{student.Phone3}  {student.Phone4}");var edit=MainWindow.PrimaryButton("EDIT");edit.Click+=async(_,_)=>{var w=new StudentFormWindow(student);await w.ShowDialog(this);if(w.Saved)Close();};var del=MainWindow.LightButton("DELETE");del.Click+=async(_,_)=>{if(await Ui.Confirm(this,"Delete Student",$"Move {student.Name} to Deleted Students?")){ExcelRepository.DeleteStudent(student.StudentNo);Close();}};var pdf=MainWindow.LightButton("PDF");pdf.Click+=async(_,_)=>await Pdf(false);var print=MainWindow.LightButton("PRINT");print.Click+=async(_,_)=>await Pdf(true);var close=MainWindow.SecondaryButton("CLOSE");close.Click+=(_,_)=>Close();var root=new Grid{Margin=new Thickness(22),RowDefinitions=RowDefinitions.Parse("Auto,Auto,*,Auto")};root.Children.Add(MainWindow.Header("STUDENT PROFILE"));var dcard=MainWindow.Card(details,new Thickness(18),"#FFFFFF");dcard.Margin=new Thickness(0,12,0,12);dcard.SetValue(Grid.RowProperty,1);root.Children.Add(dcard);var act=new StackPanel{Spacing=8,Children={new TextBlock{Text="ACTIVITY",FontWeight=FontWeight.Bold,FontSize=16},activity}};act.SetValue(Grid.RowProperty,2);root.Children.Add(act);var footer=MainWindow.Row(edit,del,print,pdf,close);footer.Margin=new Thickness(0,12,0,0);footer.SetValue(Grid.RowProperty,3);root.Children.Add(footer);return root;}
    private void AddD(Grid g,int r,int c,string l,string v){var label=new TextBlock{Text=l,FontWeight=FontWeight.Bold,Foreground=MainWindow.Brush("#475569"),Margin=new Thickness(4,8)};label.SetValue(Grid.RowProperty,r);label.SetValue(Grid.ColumnProperty,c*2);var val=new TextBlock{Text=v,Margin=new Thickness(4,8),TextWrapping=TextWrapping.Wrap};val.SetValue(Grid.RowProperty,r);val.SetValue(Grid.ColumnProperty,c*2+1);g.Children.Add(label);g.Children.Add(val);}
    private void LoadActivity(){try{var t=ActivityLogRepository.LoadActivityLog();activity.ItemsSource=t.Rows.Cast<DataRow>().Where(r=>(r.Table.Columns.Contains("Student No")?Convert.ToString(r["Student No"]):"")?.Trim()==student.StudentNo).Select(ActivityRecord.From).ToList();}catch{activity.ItemsSource=Array.Empty<ActivityRecord>();}}
    private async Task Pdf(bool open){try{var path=MainWindow.ExportPath($"Student_Profile_{student.StudentNo}","pdf");ReportExportService.ExportStudentProfilePdf(path,ExcelRepository.GetStudentByNo(student.StudentNo),ActivityLogRepository.LoadActivityLog());if(open)await MainWindow.OpenPdfAfterExport(this,path);else await Ui.Alert(this,"PDF Export Complete",path);}catch(Exception ex){await Ui.Alert(this,"PDF Error",ex.Message);}}
}

public sealed class DeletedStudentsWindow : Window
{
    private readonly ObservableCollection<StudentRecord> shown=new();private List<StudentRecord> all=new();private readonly DataGrid grid=MainWindow.GridBase();private readonly TextBox search=MainWindow.Field("Search deleted students...");
    public DeletedStudentsWindow(){Title="Deleted Students";Width=1320;Height=760;WindowStartupLocation=WindowStartupLocation.CenterOwner;MainWindow.AddStudentColumns(grid);grid.ItemsSource=shown;Content=Build();Load();}
    private Control Build(){search.PropertyChanged+=(_,e)=>{if(e.Property==TextBox.TextProperty)Apply();};var restore=MainWindow.PrimaryButton("RESTORE");restore.Click+=async(_,_)=>{if(grid.SelectedItem is not StudentRecord s){await Ui.Alert(this,"Restore","Select a student first.");return;}try{ExcelRepository.RestoreDeletedStudent(s.StudentNo);Load();}catch(Exception ex){await Ui.Alert(this,"Restore Error",ex.Message);}};var close=MainWindow.SecondaryButton("CLOSE");close.Click+=(_,_)=>Close();var root=new Grid{Margin=new Thickness(22),RowDefinitions=RowDefinitions.Parse("Auto,*")};var h=new StackPanel{Spacing=10,Children={MainWindow.Header("DELETED STUDENTS"),search,MainWindow.Row(restore,close)}};root.Children.Add(h);var host=new Border{Margin=new Thickness(0,12,0,0),Child=grid};host.SetValue(Grid.RowProperty,1);root.Children.Add(host);return root;}
    private void Load(){try{all=ExcelRepository.LoadDeletedStudents().Rows.Cast<DataRow>().Select(StudentRecord.From).ToList();Apply();}catch(Exception ex){_=Ui.Alert(this,"Database Error",ex.Message);}}
    private void Apply(){var q=(search.Text??"").Trim();shown.Clear();foreach(var s in all.Where(s=>q.Length==0||$"{s.StudentNo} {s.Name} {s.Grade} {s.Room} {s.Iquama}".Contains(q,StringComparison.OrdinalIgnoreCase)))shown.Add(s);}
}

public sealed class ActivityLogWindow : Window
{
    private readonly DataGrid grid=MainWindow.GridBase();private readonly TextBox search=MainWindow.Field("Search action, student, details...");private readonly ComboBox action=new(){MinWidth=180};private DataTable all=new();
    public ActivityLogWindow(){Title="Activity Log";Width=1380;Height=800;WindowStartupLocation=WindowStartupLocation.CenterOwner;MainWindow.Add(grid,"Date / Time","DateTimeText",170);MainWindow.Add(grid,"Action","Action",150);MainWindow.Add(grid,"Student No","StudentNo",110);MainWindow.Add(grid,"Student Name","StudentName",220);MainWindow.Add(grid,"Details","Details",620);Content=Build();Load();}
    private Control Build(){search.PropertyChanged+=(_,e)=>{if(e.Property==TextBox.TextProperty)Apply();};action.SelectionChanged+=(_,_)=>Apply();var xls=MainWindow.LightButton("EXPORT EXCEL");xls.Click+=async(_,_)=>await Export("xlsx");var pdf=MainWindow.LightButton("EXPORT PDF");pdf.Click+=async(_,_)=>await Export("pdf");var refresh=MainWindow.PrimaryButton("REFRESH");refresh.Click+=(_,_)=>Load();var close=MainWindow.SecondaryButton("CLOSE");close.Click+=(_,_)=>Close();var root=new Grid{Margin=new Thickness(22),RowDefinitions=RowDefinitions.Parse("Auto,*")};var top=new StackPanel{Spacing=10,Children={MainWindow.Header("ACTIVITY LOG"),MainWindow.Row(search,action,refresh,xls,pdf,close)}};root.Children.Add(top);var host=new Border{Margin=new Thickness(0,12,0,0),Child=grid};host.SetValue(Grid.RowProperty,1);root.Children.Add(host);return root;}
    private void Load(){try{all=ActivityLogRepository.LoadActivityLog();var acts=new[]{"All"}.Concat(all.Rows.Cast<DataRow>().Select(r=>Convert.ToString(r["Action"])?.Trim()??"").Where(x=>x.Length>0).Distinct().OrderBy(x=>x)).ToList();action.ItemsSource=acts;action.SelectedIndex=0;Apply();}catch(Exception ex){_=Ui.Alert(this,"Activity Log",ex.Message);}}
    private void Apply(){var q=(search.Text??"").Trim();var a=action.SelectedItem?.ToString()??"All";grid.ItemsSource=all.Rows.Cast<DataRow>().Select(ActivityRecord.From).Where(x=>(a=="All"||x.Action.Equals(a,StringComparison.OrdinalIgnoreCase))&&(q.Length==0||$"{x.DateTimeText} {x.Action} {x.StudentNo} {x.StudentName} {x.Details}".Contains(q,StringComparison.OrdinalIgnoreCase))).ToList();}
    private async Task Export(string ext){try{var p=MainWindow.ExportPath("Activity_Log",ext);if(ext=="xlsx")ReportExportService.ExportActivityExcel(p,all,$"Activities: {all.Rows.Count:#,0}");else ReportExportService.ExportActivityPdf(p,all,$"Activities: {all.Rows.Count:#,0}");await Ui.Alert(this,"Export Complete",p);}catch(Exception ex){await Ui.Alert(this,"Export Error",ex.Message);}}
}

public sealed class DataQualityWindow : Window
{
    private readonly DataGrid grid=MainWindow.GridBase();private DataTable table=new();
    public DataQualityWindow(){Title="Data Quality";Width=1280;Height=760;WindowStartupLocation=WindowStartupLocation.CenterOwner;Content=Build();Load();}
    private Control Build(){var refresh=MainWindow.PrimaryButton("REFRESH");refresh.Click+=(_,_)=>Load();var xls=MainWindow.LightButton("EXPORT EXCEL");xls.Click+=async(_,_)=>await Export("xlsx");var pdf=MainWindow.LightButton("EXPORT PDF");pdf.Click+=async(_,_)=>await Export("pdf");var close=MainWindow.SecondaryButton("CLOSE");close.Click+=(_,_)=>Close();var root=new Grid{Margin=new Thickness(22),RowDefinitions=RowDefinitions.Parse("Auto,*")};root.Children.Add(new StackPanel{Spacing=10,Children={MainWindow.Header("DATA QUALITY"),MainWindow.Row(refresh,xls,pdf,close)}});var h=new Border{Margin=new Thickness(0,12,0,0),Child=grid};h.SetValue(Grid.RowProperty,1);root.Children.Add(h);return root;}
    private void Load(){try{table=DataQualityScanner.ScanStudents();grid.Columns.Clear();foreach(DataColumn c in table.Columns)MainWindow.Add(grid,c.ColumnName,c.ColumnName,Math.Clamp(c.ColumnName.Length*14,100,300));grid.ItemsSource=table.DefaultView;}catch(Exception ex){_=Ui.Alert(this,"Data Quality",ex.Message);}}
    private async Task Export(string ext){try{var p=MainWindow.ExportPath("Data_Quality",ext);if(ext=="xlsx")ReportExportService.ExportDataQualityExcel(p,table,$"Issues: {table.Rows.Count:#,0}");else ReportExportService.ExportDataQualityPdf(p,table,$"Issues: {table.Rows.Count:#,0}");await Ui.Alert(this,"Export Complete",p);}catch(Exception ex){await Ui.Alert(this,"Export Error",ex.Message);}}
}

public sealed class ImportWizardWindow : Window
{
    private readonly TextBox file=MainWindow.Field("Excel workbook path");private readonly DataGrid preview=MainWindow.GridBase();private DataTable? data;
    public ImportWizardWindow(){Title="Import Students";Width=1200;Height=760;WindowStartupLocation=WindowStartupLocation.CenterOwner;Content=Build();}
    private Control Build(){var browse=MainWindow.PrimaryButton("SELECT EXCEL");browse.Click+=async(_,_)=>await Browse();var load=MainWindow.LightButton("PREVIEW");load.Click+=async(_,_)=>await LoadPreview(file.Text??"");var import=MainWindow.PrimaryButton("IMPORT STUDENTS");import.Click+=async(_,_)=>await Import();var close=MainWindow.SecondaryButton("CLOSE");close.Click+=(_,_)=>Close();var root=new Grid{Margin=new Thickness(22),RowDefinitions=RowDefinitions.Parse("Auto,*")};root.Children.Add(new StackPanel{Spacing=10,Children={MainWindow.Header("IMPORT STUDENTS"),new TextBlock{Text="Select an Excel workbook, review the rows, then import.",Foreground=MainWindow.Brush("#64748B")},MainWindow.Row(file,browse,load,import,close)}});var h=new Border{Margin=new Thickness(0,12,0,0),Child=preview};h.SetValue(Grid.RowProperty,1);root.Children.Add(h);return root;}
    private async Task Browse(){try{var fs=await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions{Title="Import Students from Excel",AllowMultiple=false,FileTypeFilter=new[]{new FilePickerFileType("Excel workbook"){Patterns=new[]{"*.xlsx","*.xlsm"}}}});var f=fs.FirstOrDefault();if(f is null)return;var p=f.TryGetLocalPath();if(string.IsNullOrWhiteSpace(p))return;file.Text=p;await LoadPreview(p);}catch(Exception ex){await Ui.Alert(this,"Import",ex.Message);}}
    private async Task LoadPreview(string path){try{if(!File.Exists(path))throw new Exception("Excel file not found.");data=ReadWorkbook(path);preview.Columns.Clear();foreach(DataColumn c in data.Columns)MainWindow.Add(preview,c.ColumnName,c.ColumnName,Math.Clamp(c.ColumnName.Length*14,90,180));preview.ItemsSource=data.DefaultView;}catch(Exception ex){await Ui.Alert(this,"Preview Error",ex.Message);}}
    private async Task Import(){try{if(data is null||data.Rows.Count==0)throw new Exception("Load a workbook preview first.");var count=ExcelRepository.ImportStudentsBatch(data);await Ui.Alert(this,"Import Complete",$"{count:#,0} students imported successfully.");Close();}catch(Exception ex){await Ui.Alert(this,"Import Error",ex.Message);}}
    private static DataTable ReadWorkbook(string path){string[] fields={"Student No","Name","Gender","Iquama","Date of Birth","Grade","Room","Phone No","Phone No2","Phone No3","Phone No4"};var dt=new DataTable();foreach(var f in fields)dt.Columns.Add(f);using var wb=new XLWorkbook(path);var ws=wb.Worksheets.First();var headers=new Dictionary<string,int>(StringComparer.OrdinalIgnoreCase);var last=ws.LastColumnUsed()?.ColumnNumber()??0;for(int c=1;c<=last;c++){var h=ws.Cell(1,c).GetFormattedString().Trim();if(h.Length>0&&!headers.ContainsKey(h))headers[h]=c;}if(!headers.ContainsKey("Student No")||!headers.ContainsKey("Name"))throw new Exception("Workbook must contain Student No and Name headers in row 1.");var end=ws.LastRowUsed()?.RowNumber()??1;for(int r=2;r<=end;r++){var row=dt.NewRow();bool any=false;foreach(var f in fields){var v=headers.TryGetValue(f,out var c)?ws.Cell(r,c).GetFormattedString().Trim():"";row[f]=v;if(v.Length>0)any=true;}if(any)dt.Rows.Add(row);}return dt;}
}

public sealed class SettingsWindow : Window
{
    private readonly TextBox school=MainWindow.Field(),db=MainWindow.Field(),export=MainWindow.Field();private readonly CheckBox backup=new(){Content="Automatic backup enabled"};private readonly NumericUpDown retention=new(){Minimum=1,Maximum=365,Width=130};
    public SettingsWindow(){Title="Settings";Width=980;Height=820;WindowStartupLocation=WindowStartupLocation.CenterOwner;Content=Build();Load();}
    private Control Build(){var save=MainWindow.PrimaryButton("SAVE");save.Click+=async(_,_)=>await Save();var reset=MainWindow.LightButton("RESET DEFAULTS");reset.Click+=async(_,_)=>{if(await Ui.Confirm(this,"Reset Settings","Reset application settings to defaults?")){AppSettingsManager.ResetToDefaults();Load();}};var pass=MainWindow.LightButton("CHANGE PASSWORD");pass.Click+=async(_,_)=>await new ChangePasswordWindow().ShowDialog(this);var close=MainWindow.SecondaryButton("CLOSE");close.Click+=(_,_)=>Close();var content=new StackPanel{Spacing=12,MaxWidth=820,Children={MainWindow.Header("SETTINGS"),new TextBlock{Text="APPLICATION",FontWeight=FontWeight.Bold},new TextBlock{Text="School name"},school,new TextBlock{Text="DATABASE",FontWeight=FontWeight.Bold},new TextBlock{Text="Database folder (Mac default: Documents/StudentsDatabase)"},db,new TextBlock{Text="BACKUP",FontWeight=FontWeight.Bold},backup,new TextBlock{Text="Backup retention count"},retention,new TextBlock{Text="EXPORT",FontWeight=FontWeight.Bold},new TextBlock{Text="Default export folder"},export,new TextBlock{Text="ADMIN SECURITY",FontWeight=FontWeight.Bold},new TextBlock{Text=$"Administrator: {AdminAuthManager.GetAdminUsername()}"},MainWindow.Row(save,reset,pass,close)}};return new ScrollViewer{Content=new Border{Margin=new Thickness(28),Child=MainWindow.Card(content,new Thickness(24),"#FFFFFF")}};}
    private void Load(){var s=AppSettingsManager.Current;school.Text=s.SchoolName;db.Text=s.DatabaseFolder;export.Text=s.DefaultExportFolder;backup.IsChecked=s.AutomaticBackupEnabled;retention.Value=s.BackupRetentionCount;}
    private async Task Save(){try{AppSettingsManager.UpdateSettings(school.Text??"",db.Text??"",backup.IsChecked==true,(int)(retention.Value??30),export.Text??"");await Ui.Alert(this,"Settings","Settings saved successfully.");}catch(Exception ex){await Ui.Alert(this,"Settings Error",ex.Message);}}
}