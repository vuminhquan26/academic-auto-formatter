using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Windows.ApplicationModel.DataTransfer;
using Windows.Storage.Pickers;

namespace AutoFormatter
{
    public sealed partial class MainWindow : Window
    {
        private string _targetFilePath = "";

        public MainWindow() { this.InitializeComponent(); }

        private void DropZone_DragOver(object sender, DragEventArgs e)
        {
            e.AcceptedOperation = DataPackageOperation.Copy;
            e.DragUIOverride.Caption = "Thả file Word báo cáo (.docx) vào đây";
        }

        private async void DropZone_Drop(object sender, DragEventArgs e)
        {
            if (e.DataView.Contains(StandardDataFormats.StorageItems))
            {
                var items = await e.DataView.GetStorageItemsAsync();
                var file = items.FirstOrDefault();
                if (file != null && file.Path.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                {
                    _targetFilePath = file.Path;
                    TxtFileName.Text = $"Đã chọn: {file.Name}";
                    TxtStatus.Text = "Trạng thái: Sẵn sàng.";
                }
            }
        }

        private async void BtnBrowseLogo_Click(object sender, RoutedEventArgs e)
        {
            var picker = new FileOpenPicker();
            picker.ViewMode = PickerViewMode.List;
            picker.SuggestedStartLocation = PickerLocationId.PicturesLibrary;
            picker.FileTypeFilter.Add(".png");
            picker.FileTypeFilter.Add(".jpg");

            var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            WinRT.Interop.InitializeWithWindow.Initialize(picker, hwnd);

            var file = await picker.PickSingleFileAsync();
            if (file != null) TxtLogoPath.Text = file.Path;
        }

        private void TgIsGroup_Toggled(object sender, RoutedEventArgs e)
        {
            TxtGroupName.Visibility = TgIsGroup.IsOn ? Visibility.Visible : Visibility.Collapsed;
        }

        private async void BtnStart_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_targetFilePath))
            {
                TxtStatus.Text = "Lỗi: Vui lòng kéo thả file báo cáo (.docx) vào trước!";
                return;
            }

            var config = new FormatConfig
            {
                PageSetup = new PageSetupConfig { LeftCm = NumLeft.Value, RightCm = NumRight.Value, TopCm = NumTop.Value, BottomCm = NumBottom.Value, IsTwoSidedPrint = TgTwoSided.IsOn },
                Typography = new TypographyConfig { FontName = CboFont.SelectedItem.ToString(), FontSize = (int)NumFontSize.Value, EnforceLineSpacing = TgSpacing.IsOn, LineSpacingValue = NumLineSpacing.Value, AutoFormatHeadings = TgHeadings.IsOn },
                Automation = new AutomationConfig { FixTableAndFigureWidth = TgTableFigure.IsOn, FixEquations = TgEquation.IsOn, FixReferenceIndentation = TgReference.IsOn },
                Toc = new TocConfig { GenerateTOC = TgTOC.IsOn, GenerateTOF = TgTOF.IsOn, GenerateTOT = TgTOT.IsOn },
                Cover = new CoverGeneratorConfig
                {
                    GenerateCover = TgGenCover.IsOn,
                    AddPageBorder = TgPageBorder.IsOn,
                    UniversityName = TxtUni.Text,
                    FacultyName = TxtFaculty.Text,
                    ReportType = TxtReportType.Text,
                    TopicName = TxtTopic.Text,
                    LogoPath = TxtLogoPath.Text,
                    InstructorName = TxtInstructor.Text,
                    LocationAndDate = TxtLocation.Text,
                    IsGroupProject = TgIsGroup.IsOn,
                    GroupName = TxtGroupName.Text
                }
            };

            var rawCommon = TxtCommonFields.Text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in rawCommon)
            {
                var parts = line.Split(new[] { ':' }, 2, StringSplitOptions.TrimEntries);
                if (parts.Length == 2) config.Cover.CommonFields.Add(new CustomField { Label = parts[0] + ":", Value = parts[1] });
                else config.Cover.CommonFields.Add(new CustomField { Label = "", Value = line });
            }

            string[] rawStudents = TxtStudentsData.Text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in rawStudents)
            {
                var parts = line.Split(new[] { '-' }, StringSplitOptions.TrimEntries);
                if (parts.Length >= 1)
                {
                    var stu = new StudentInfo { Name = parts[0] };
                    if (parts.Length >= 2) stu.StudentId = parts[1];
                    config.Cover.Students.Add(stu);
                }
            }

            ProgBar.Visibility = Visibility.Visible;
            BtnStart.IsEnabled = false;
            TxtStatus.Text = "Trạng thái: Đang xử lý...";

            string directory = Path.GetDirectoryName(_targetFilePath);
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(_targetFilePath);
            string extension = Path.GetExtension(_targetFilePath);
            string outputPath = Path.Combine(directory, $"{fileNameWithoutExt}_HoanThien{extension}");

            await Task.Run(() =>
            {
                try
                {
                    File.Copy(_targetFilePath, outputPath, true);
                    DocxEngine.FormatDocument(outputPath, config);
                }
                catch (Exception ex)
                {
                    DispatcherQueue.TryEnqueue(() => TxtStatus.Text = $"Lỗi: {ex.Message}");
                }
            });

            ProgBar.Visibility = Visibility.Collapsed;
            BtnStart.IsEnabled = true;
            TxtStatus.Text = $"Hoàn tất! File đã lưu tại: {Path.GetFileName(outputPath)}";

            try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputPath) { UseShellExecute = true }); } catch { }
        }
    }
}