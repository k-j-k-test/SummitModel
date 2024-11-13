using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ActuLiteModel;
using Newtonsoft.Json;
using LiveCharts.Wpf;
using LiveCharts;
using System.Text;
using System.Threading;
using System.Collections;
using System.Windows.Data;
using System.ComponentModel;
using System.Globalization;
using System.Windows.Threading;
using System.Windows.Media;
using LiveCharts.Wpf.Charts.Base;
using ModernWpf.Controls;
using System.Diagnostics;
using System.IO;

namespace ActuLight
{
    public partial class RiskRateManagementWindow : Window
    {
        private readonly ApiClient _apiClient;
        private int currentPage = 1;
        public int PageSize = 1000;
        private string lastSearchTerm = "";
        private Dictionary<string, RiskRateSummary> summaryCache = new Dictionary<string, RiskRateSummary>();
        private readonly RiskRateChartManager _chartManager;
        private string selectedCompany = "전체 회사";
        private List<string> companies = new List<string>();

        public RiskRateManagementWindow()
        {
            InitializeComponent();
            InitializeGridOptimizations();
            _apiClient = new ApiClient();
            _chartManager = new RiskRateChartManager(AnalysisChart);
            SetupDetailGridColumns();
            InitializeAnalysisTabContextMenu();
            InitializeSearchGridContextMenu();
            InitializeCompanyComboBox();
        }

        private void InitializeGridOptimizations()
        {
            // 스크롤 성능 최적화를 위한 설정
            DetailResultsGrid.EnableColumnVirtualization = true;
            DetailResultsGrid.EnableRowVirtualization = true;

            // 가상화 모드 설정
            VirtualizingPanel.SetIsVirtualizing(DetailResultsGrid, true);
            VirtualizingPanel.SetVirtualizationMode(DetailResultsGrid, VirtualizationMode.Recycling);

            // ScrollUnit을 Pixel로 설정하여 스크롤 성능 향상
            ScrollViewer.SetCanContentScroll(DetailResultsGrid, true);
            VirtualizingPanel.SetScrollUnit(DetailResultsGrid, ScrollUnit.Pixel);

            // 캐시 길이 최적화
            VirtualizingPanel.SetCacheLength(DetailResultsGrid, new VirtualizationCacheLength(5));
            VirtualizingPanel.SetCacheLengthUnit(DetailResultsGrid, VirtualizationCacheLengthUnit.Page);

            // UI 업데이트 지연을 줄이기 위한 설정
            DetailResultsGrid.FrozenColumnCount = 13; // 기본 정보 컬럼을 고정

            // 렌더링 최적화
            RenderOptions.SetBitmapScalingMode(DetailResultsGrid, BitmapScalingMode.LowQuality);
            RenderOptions.SetEdgeMode(DetailResultsGrid, EdgeMode.Aliased);

            // 불필요한 기능 비활성화
            DetailResultsGrid.CanUserReorderColumns = false;
            DetailResultsGrid.CanUserSortColumns = false;
            DetailResultsGrid.CanUserResizeColumns = false;
        }


        // 검색 탭 관련 메서드

        private async void SearchButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                currentPage = 1; // 새 검색시 첫 페이지로 리셋
                lastSearchTerm = SearchBox.Text?.Trim() ?? string.Empty; // 빈 문자열은 ApiClient에서 "empty"로 변환됨
                await PerformSearch();
            }
            catch (ApiException ex)
            {
                MessageBox.Show($"검색 중 오류가 발생했습니다. {ex.ResponseContent}",
                    "검색 오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"예상치 못한 오류가 발생했습니다. {ex.Message}",
                    "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task PerformSearch()
        {
            var searchResult = await _apiClient.SearchRiskRatesAsync(
                lastSearchTerm,
                currentPage,
                PageSize,
                selectedCompany == "전체 회사" ? null : selectedCompany);

            // 검색 결과를 RiskRateSummary로 변환하고 그룹화
            var summaries = searchResult.Items
                .GroupBy(r =>
                {
                    string[] factors = r.FACTORS.Split('|');
                    var factorString = string.Join(" ", Enumerable.Range(1, factors.Length - 1)
                            .Where(i => !string.IsNullOrWhiteSpace(factors[i]))
                            .Select(i => $"F{i}"));

                    return new
                    {
                        r.COMPANY_NAME,
                        r.RISK_RATE_NAME,
                        r.APPROVAL_DATE,
                        기간구조 = factors[0],
                        팩터구성 = factorString
                    };
                })
                .Select(g =>
                {
                    var firstItem = g.First();
                    var factors = firstItem.FACTORS.Split('|');

                    var factorIndices = Enumerable.Range(1, factors.Length - 1)
                        .Where(i => !string.IsNullOrWhiteSpace(factors[i]))
                        .ToList();

                    var summary = RiskRateSummary.FromRiskRate(firstItem);
                    summary.Details = g.OrderBy(r =>
                    {
                        var itemFactors = r.FACTORS.Split('|');
                        return factorIndices.Select(i =>
                        {
                            var factorValue = itemFactors[i];
                            return int.TryParse(factorValue, out int value) ? value : 0;
                        }).ToList();
                    }, new FactorComparer()).ToList();

                    return summary;
                })
                .ToList();

            // 캐시 업데이트
            summaryCache = summaries.ToDictionary(s => s.Key);

            SearchResultsGrid.ItemsSource = summaries;
            ResultCountText.Text = $"검색 결과: {summaries.Count}건";
        }

        private void SearchResultsGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var selectedSummaries = SearchResultsGrid.SelectedItems.Cast<RiskRateSummary>().ToList();
            if (selectedSummaries.Any())
            {
                ShowRiskRateDetails(selectedSummaries);
            }
        }

        private void SetupDetailGridColumns()
        {
            DetailResultsGrid.Columns.Clear();

            // 기본 컬럼 추가
            DetailResultsGrid.Columns.Add(new DataGridTextColumn { Header = "회사명", Binding = new Binding("COMPANY_NAME"), Width = new DataGridLength(100) });
            DetailResultsGrid.Columns.Add(new DataGridTextColumn { Header = "위험률명", Binding = new Binding("RISK_RATE_NAME"), Width = new DataGridLength(150) });
            DetailResultsGrid.Columns.Add(new DataGridTextColumn { Header = "적용일자", Binding = new Binding("APPROVAL_DATE"), Width = new DataGridLength(80) });
            DetailResultsGrid.Columns.Add(new DataGridTextColumn { Header = "기간구분", Binding = new Binding("기간구분"), Width = new DataGridLength(80) });

            // 팩터 컬럼 (F1-F9)
            for (int i = 1; i <= 9; i++)
            {
                DetailResultsGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = $"F{i}",
                    Binding = new Binding($"F{i}"),
                    Width = new DataGridLength(60)
                });
            }

            // 위험률 컬럼 - 배열 인덱스로 직접 접근
            for (int i = 0; i < 241; i++)
            {
                DetailResultsGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = i.ToString(),
                    Binding = new Binding($"RateValues[{i}]"),
                    Width = new DataGridLength(60)
                });
            }

            // 비고 컬럼
            DetailResultsGrid.Columns.Add(new DataGridTextColumn { Header = "비고", Binding = new Binding("COMMENT"), Width = new DataGridLength(150) });
        }

        private void InitializeSearchGridContextMenu()
        {
            SearchResultsGrid.ContextMenu = new ContextMenu();

            // 상세보기 메뉴 아이템
            var viewDetailsMenuItem = new MenuItem
            {
                Header = "상세보기",
            };
            viewDetailsMenuItem.Click += ViewDetails_Click;

            // 상세보기(추가) 메뉴 아이템
            var viewDetailsAddMenuItem = new MenuItem
            {
                Header = "상세보기(추가)",
            };
            viewDetailsAddMenuItem.Click += ViewDetailsAdd_Click;

            // 일반 복사 메뉴 아이템
            var copyMenuItem = new MenuItem
            {
                Header = "복사",
            };
            copyMenuItem.Click += CopyRiskRate_Click;

            // PVPLUS 형식 복사 메뉴 아이템
            var copyPvPlusMenuItem = new MenuItem
            {
                Header = "복사(PVPLUS)",
            };
            copyPvPlusMenuItem.Click += CopyRiskRatePvPlus_Click;

            // 이름 변경 메뉴 아이템
            var renameMenuItem = new MenuItem 
            {
                Header = "이름 변경" 
            };
            renameMenuItem.Click += RenameRiskRate_Click;

            // 삭제 메뉴 아이템
            var deleteMenuItem = new MenuItem
            {
                Header = "위험률 삭제",
            };
            deleteMenuItem.Click += DeleteRiskRate_Click;

            SearchResultsGrid.ContextMenu.Items.Add(viewDetailsMenuItem);
            SearchResultsGrid.ContextMenu.Items.Add(viewDetailsAddMenuItem);
            SearchResultsGrid.ContextMenu.Items.Add(copyMenuItem);
            SearchResultsGrid.ContextMenu.Items.Add(copyPvPlusMenuItem);
            SearchResultsGrid.ContextMenu.Items.Add(renameMenuItem);
            SearchResultsGrid.ContextMenu.Items.Add(deleteMenuItem);

            // 컨텍스트 메뉴가 열릴 때 선택된 항목 확인
            SearchResultsGrid.ContextMenu.Opened += (s, e) =>
            {
                var selectedItems = SearchResultsGrid.SelectedItems.Cast<RiskRateSummary>().ToList();
                var isItemSelected = selectedItems.Any();
                var isSingleItemSelected = selectedItems.Count == 1;

                viewDetailsMenuItem.IsEnabled = isItemSelected;
                viewDetailsAddMenuItem.IsEnabled = isItemSelected;
                copyMenuItem.IsEnabled = isItemSelected;
                copyPvPlusMenuItem.IsEnabled = isItemSelected;
                renameMenuItem.IsEnabled = isSingleItemSelected;
                deleteMenuItem.IsEnabled = isItemSelected;

                // 선택된 항목 수에 따라 메뉴 텍스트 업데이트
                if (SearchResultsGrid.SelectedItems.Count > 1)
                {
                    viewDetailsMenuItem.Header = $"상세보기 ({SearchResultsGrid.SelectedItems.Count}개 항목)";
                    viewDetailsAddMenuItem.Header = $"상세보기(추가) ({SearchResultsGrid.SelectedItems.Count}개 항목)";
                }
                else
                {
                    viewDetailsMenuItem.Header = "상세보기";
                    viewDetailsAddMenuItem.Header = "상세보기(추가)";
                }
            };
        }

        private async void InitializeCompanyComboBox()
        {
            try
            {
                // 회사 목록 가져오기
                companies = await _apiClient.GetCompanyNamesAsync();

                // "전체 회사" 옵션 추가
                companies.Insert(0, "전체 회사");

                // ComboBox 데이터 설정
                CompanyComboBox.ItemsSource = companies;
                CompanyComboBox.SelectedItem = "전체 회사";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"회사 목록을 가져오는 중 오류가 발생했습니다.\n{ex.Message}",
                    "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void CompanyComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CompanyComboBox.SelectedItem is string company)
            {
                selectedCompany = company;

                // 검색어가 있는 경우 자동으로 재검색
                if (!string.IsNullOrWhiteSpace(SearchBox.Text))
                {
                    await PerformSearch();
                }
            }
        }

        private void ViewDetails_Click(object sender, RoutedEventArgs e)
        {
            var selectedSummaries = SearchResultsGrid.SelectedItems.Cast<RiskRateSummary>().ToList();
            if (selectedSummaries.Any())
            {
                ShowRiskRateDetails(selectedSummaries);
            }
        }

        private void ViewDetailsAdd_Click(object sender, RoutedEventArgs e)
        {
            var selectedSummaries = SearchResultsGrid.SelectedItems.Cast<RiskRateSummary>().ToList();
            if (selectedSummaries.Any())
            {
                AppendRiskRateDetails(selectedSummaries);
            }
        }

        private void AppendRiskRateDetails(List<RiskRateSummary> selectedSummaries)
        {
            // 현재 DetailResultsGrid의 아이템들을 가져옴
            var currentDetails = DetailResultsGrid.ItemsSource?.Cast<RiskRateDetailViewModel>().ToList()
                                ?? new List<RiskRateDetailViewModel>();

            // 새로운 상세 정보를 생성
            var newDetails = selectedSummaries
                .SelectMany(summary => summary.Details)
                .Select(rr =>
                {
                    var factors = rr.FACTORS.Split('|');
                    var rates = (rr.RATES ?? "").Split('|');

                    return new RiskRateDetailViewModel
                    {
                        COMPANY_NAME = rr.COMPANY_NAME,
                        RISK_RATE_NAME = rr.RISK_RATE_NAME,
                        APPROVAL_DATE = rr.APPROVAL_DATE,
                        FACTORS = rr.FACTORS,
                        RATES = rr.RATES,
                        COMMENT = rr.COMMENT,
                        기간구분 = factors.Length > 0 ? factors[0] : "",
                        F1 = factors.Length > 1 ? factors[1] : "",
                        F2 = factors.Length > 2 ? factors[2] : "",
                        F3 = factors.Length > 3 ? factors[3] : "",
                        F4 = factors.Length > 4 ? factors[4] : "",
                        F5 = factors.Length > 5 ? factors[5] : "",
                        F6 = factors.Length > 6 ? factors[6] : "",
                        F7 = factors.Length > 7 ? factors[7] : "",
                        F8 = factors.Length > 8 ? factors[8] : "",
                        F9 = factors.Length > 9 ? factors[9] : "",
                        RateValues = rates
                    };
                });

            // 기존 목록과 새로운 목록을 합침
            var combinedDetails = currentDetails.Concat(newDetails).ToList();

            // 중복 제거 (모든 필드가 동일한 항목 제거)
            var distinctDetails = combinedDetails
                .GroupBy(x => $"{x.COMPANY_NAME}|{x.RISK_RATE_NAME}|{x.APPROVAL_DATE}|{x.FACTORS}")
                .Select(g => g.First())
                .ToList();

            DetailResultsGrid.ItemsSource = distinctDetails;

            // 상세 정보 탭으로 전환
            MainTabControl.SelectedItem = DetailTab;

            // 선택된 위험률 정보 텍스트 업데이트
            var totalCount = distinctDetails.Count;
            var newCount = selectedSummaries.Sum(s => s.Details.Count);
            SelectedRiskRateInfo.Text = $"선택된 위험률: {totalCount}개 (신규 {newCount}개 추가)";

            // DataGrid가 렌더링될 때까지 기다린 후 컬럼 크기 자동 조정
            Dispatcher.BeginInvoke(new Action(() =>
            {
                foreach (var column in DetailResultsGrid.Columns)
                {
                    if (column.Header.ToString() == "위험률명") continue;
                    column.Width = new DataGridLength(1, DataGridLengthUnitType.Auto);
                }
            }), DispatcherPriority.Loaded);
        }

        private void CopyRiskRate_Click(object sender, RoutedEventArgs e)
        {
            var selectedSummaries = SearchResultsGrid.SelectedItems.Cast<RiskRateSummary>().ToList();
            if (selectedSummaries.Any())
            {
                try
                {
                    var allLines = selectedSummaries.SelectMany(summary =>
                        summary.Details.Select(detail =>
                            string.Join("\t",
                                detail.COMPANY_NAME,
                                detail.RISK_RATE_NAME,
                                detail.APPROVAL_DATE,
                                detail.FACTORS,
                                detail.RATES,
                                detail.COMMENT ?? ""
                            )
                        ));

                    var textToCopy = string.Join(Environment.NewLine, allLines);
                    Clipboard.SetText(textToCopy);

                    MessageBox.Show($"{selectedSummaries.Count}개의 위험률이 클립보드에 복사되었습니다.",
                        "복사 완료", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"복사 중 오류가 발생했습니다.\n{ex.Message}",
                        "복사 오류", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void CopyRiskRatePvPlus_Click(object sender, RoutedEventArgs e)
        {
            var selectedSummaries = SearchResultsGrid.SelectedItems.Cast<RiskRateSummary>().ToList();
            if (selectedSummaries.Any())
            {
                try
                {
                    var lines = new List<string>();

                    foreach (var summary in selectedSummaries)
                    {
                        foreach (var detail in summary.Details)
                        {
                            var factors = detail.FACTORS.Split('|');
                            var rates = detail.RATES.Split('|');

                            var fields = new List<string>
                    {
                        detail.RISK_RATE_NAME,
                        detail.APPROVAL_DATE,
                        factors[0]  // 기간구분
                    };

                            // F1~F9 추가
                            for (int i = 1; i <= 9; i++)
                            {
                                fields.Add(factors.Length > i ? factors[i] : "");
                            }

                            // RateValues 추가
                            fields.AddRange(rates);

                            lines.Add(string.Join("\t", fields));
                        }
                    }

                    var textToCopy = string.Join(Environment.NewLine, lines);
                    Clipboard.SetText(textToCopy);

                    MessageBox.Show($"{selectedSummaries.Count}개의 위험률이 PVPLUS 형식으로 클립보드에 복사되었습니다.",
                        "복사 완료", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"PVPLUS 형식 복사 중 오류가 발생했습니다.\n{ex.Message}",
                        "복사 오류", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private async void DeleteRiskRate_Click(object sender, RoutedEventArgs e)
        {
            var selectedSummaries = SearchResultsGrid.SelectedItems.Cast<RiskRateSummary>().ToList();
            if (selectedSummaries.Any())
            {
                var message = selectedSummaries.Count == 1
                    ? $"선택한 위험률을 삭제하시겠습니까?\n\n회사명: {selectedSummaries[0].회사명}\n위험률명: {selectedSummaries[0].위험률명}\n적용일자: {selectedSummaries[0].적용일자}"
                    : $"선택한 {selectedSummaries.Count}개의 위험률을 삭제하시겠습니까?";

                var result = MessageBox.Show(
                    message,
                    "위험률 삭제 확인",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning,
                    MessageBoxResult.No);

                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        var deleteTasks = selectedSummaries
                            .SelectMany(summary => summary.Details)
                            .Select(detail =>
                                _apiClient.DeleteRiskRateAsync(
                                    detail.COMPANY_NAME,
                                    detail.RISK_RATE_NAME,
                                    detail.APPROVAL_DATE,
                                    detail.FACTORS));

                        await Task.WhenAll(deleteTasks);

                        MessageBox.Show($"{selectedSummaries.Count}개의 위험률이 성공적으로 삭제되었습니다.",
                            "삭제 완료", MessageBoxButton.OK, MessageBoxImage.Information);

                        // 검색 결과 새로고침
                        await PerformSearch();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            $"위험률 삭제 중 오류가 발생했습니다.\n{ex.Message}",
                            "삭제 오류",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);
                    }
                }
            }
        }

        private void ShowRiskRateDetails(List<RiskRateSummary> selectedSummaries)
        {
            var detailViewModels = selectedSummaries
                .SelectMany(summary => summary.Details)
                .Select(rr =>
                {
                    var factors = rr.FACTORS.Split('|');
                    var rates = (rr.RATES ?? "").Split('|');

                    return new RiskRateDetailViewModel
                    {
                        COMPANY_NAME = rr.COMPANY_NAME,
                        RISK_RATE_NAME = rr.RISK_RATE_NAME,
                        APPROVAL_DATE = rr.APPROVAL_DATE,
                        FACTORS = rr.FACTORS,
                        RATES = rr.RATES,
                        COMMENT = rr.COMMENT,
                        기간구분 = factors.Length > 0 ? factors[0] : "",
                        F1 = factors.Length > 1 ? factors[1] : "",
                        F2 = factors.Length > 2 ? factors[2] : "",
                        F3 = factors.Length > 3 ? factors[3] : "",
                        F4 = factors.Length > 4 ? factors[4] : "",
                        F5 = factors.Length > 5 ? factors[5] : "",
                        F6 = factors.Length > 6 ? factors[6] : "",
                        F7 = factors.Length > 7 ? factors[7] : "",
                        F8 = factors.Length > 8 ? factors[8] : "",
                        F9 = factors.Length > 9 ? factors[9] : "",
                        RateValues = rates
                    };
                })
                .ToList();

            DetailResultsGrid.ItemsSource = detailViewModels;

            // 선택된 위험률 정보 텍스트 업데이트
            var totalCount = detailViewModels.Count;
            SelectedRiskRateInfo.Text = totalCount == 1
                ? $"선택된 위험률: {detailViewModels[0].COMPANY_NAME} - {detailViewModels[0].RISK_RATE_NAME}"
                : $"선택된 위험률: {totalCount}개";

            // DataGrid가 렌더링될 때까지 기다린 후 컬럼 크기 자동 조정
            Dispatcher.BeginInvoke(new Action(() =>
            {
                foreach (var column in DetailResultsGrid.Columns)
                {
                    if (column.Header.ToString() == "위험률명") continue;
                    column.Width = new DataGridLength(1, DataGridLengthUnitType.Auto);
                }
                MainTabControl.SelectedItem = DetailTab;
            }), DispatcherPriority.Loaded);
        }

        private void SearchBox_SuggestionChosen(AutoSuggestBox sender, AutoSuggestBoxSuggestionChosenEventArgs args)
        {
            if (args.SelectedItem is string selectedRiskRate)
            {
                sender.Text = selectedRiskRate;
            }
        }

        private async void SearchBox_QuerySubmitted(AutoSuggestBox sender, AutoSuggestBoxQuerySubmittedEventArgs args)
        {
            // 자동완성 항목을 선택했거나 검색어가 비어있는 경우 바로 검색 실행
            if (args.ChosenSuggestion != null || string.IsNullOrWhiteSpace(sender.Text))
            {
                SearchButton_Click(sender, new RoutedEventArgs());
            }
        }

        private async void SearchBox_TextChanged(AutoSuggestBox sender, AutoSuggestBoxTextChangedEventArgs args)
        {
            if (args.Reason == AutoSuggestionBoxTextChangeReason.UserInput)
            {
                // 2글자 미만이면 자동완성 제안을 표시하지 않음
                if (sender.Text.Length < 2)
                {
                    sender.ItemsSource = new List<string>();
                    return;
                }

                await DebouncerAsync.Debounce("AutoComplete", 150, async () =>
                {
                    try
                    {
                        var suggestions = await _apiClient.GetRiskRateNamesAsync(
                            sender.Text,
                            selectedCompany == "전체 회사" ? null : selectedCompany);

                        await Dispatcher.InvokeAsync(() =>
                        {
                            sender.ItemsSource = suggestions;
                        });
                    }
                    catch (Exception)
                    {
                        await Dispatcher.InvokeAsync(() =>
                        {
                            sender.ItemsSource = new List<string>();
                        });
                    }
                });
            }
        }

        // 이름 변경 관련
        private async void RenameRiskRate_Click(object sender, RoutedEventArgs e)
        {
            var selectedSummary = SearchResultsGrid.SelectedItem as RiskRateSummary;
            if (selectedSummary == null) return;

            // 사용자 입력을 위한 다이얼로그 생성
            var dialog = new ContentDialog()
            {
                Title = "위험률 이름 변경",
                PrimaryButtonText = "변경",
                CloseButtonText = "취소",
                DefaultButton = ContentDialogButton.Primary,
                Content = new TextBox
                {
                    Text = selectedSummary.위험률명,
                    Width = 300,
                    MaxLength = 100,
                    SelectionStart = 0,
                    SelectionLength = selectedSummary.위험률명.Length
                }
            };

            try
            {
                var result = await dialog.ShowAsync();
                if (result == ContentDialogResult.Primary)
                {
                    var textBox = dialog.Content as TextBox;
                    string newName = textBox.Text.Trim();

                    if (string.IsNullOrWhiteSpace(newName))
                    {
                        MessageBox.Show("새 이름을 입력해주세요.", "오류", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    if (newName == selectedSummary.위험률명)
                    {
                        return; // 이름이 변경되지 않았으면 아무 작업도 하지 않음
                    }

                    // 사용자에게 변경 범위 확인
                    var confirmMessage = $"선택한 위험률의 모든 팩터 조합({selectedSummary.Details.Count}개)에 대해 이름을 변경하시겠습니까?\n\n" +
                                       $"회사명: {selectedSummary.회사명}\n" +
                                       $"현재 이름: {selectedSummary.위험률명}\n" +
                                       $"새 이름: {newName}";

                    var confirmResult = MessageBox.Show(confirmMessage, "이름 변경 확인",
                        MessageBoxButton.YesNo, MessageBoxImage.Question);

                    if (confirmResult == MessageBoxResult.Yes)
                    {
                        try
                        {
                            var renameTasks = selectedSummary.Details.Select(detail =>
                                _apiClient.RenameRiskRateAsync(
                                    selectedSummary.회사명,
                                    selectedSummary.위험률명,
                                    detail.APPROVAL_DATE,
                                    detail.FACTORS,
                                    newName
                                ));

                            await Task.WhenAll(renameTasks);

                            MessageBox.Show(
                                $"위험률 이름이 성공적으로 변경되었습니다.\n변경된 항목 수: {selectedSummary.Details.Count}개",
                                "완료",
                                MessageBoxButton.OK,
                                MessageBoxImage.Information);

                            // 검색 결과 새로고침
                            await PerformSearch();
                        }
                        catch (ApiException ex) when (ex.StatusCode == System.Net.HttpStatusCode.Conflict)
                        {
                            MessageBox.Show(
                                "동일한 키를 가진 위험률이 이미 존재합니다.",
                                "오류",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(
                                $"위험률 이름 변경 중 오류가 발생했습니다.\n{ex.Message}",
                                "오류",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"다이얼로그 표시 중 오류가 발생했습니다.\n{ex.Message}",
                    "오류",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }



        // 등록 탭 관련 메서드
        private async void RegisterButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var inputText = RegistrationTextBox.Text.Trim();
                if (string.IsNullOrEmpty(inputText))
                {
                    MessageBox.Show("등록할 데이터를 입력해주세요.", "알림");
                    return;
                }

                var createdBy = CreatedByTextBox.Text.Trim();
                if (string.IsNullOrEmpty(createdBy))
                {
                    createdBy = "unknown";
                }

                RegisterButton.IsEnabled = false;
                ClearButton.IsEnabled = false;

                var lines = inputText.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                var remainingLines = new List<string>(lines);
                var failedRegistrations = new List<FailedRegistration>();
                var totalCount = lines.Length;

                // 마지막 라인에 탭 추가
                if (lines.Length > 0)
                {
                    lines[lines.Length - 1] = lines[lines.Length - 1] + "\t";
                }

                // 중복 제거를 위한 Dictionary (복합키를 문자열로 만들어 사용)
                var uniqueRiskRates = new Dictionary<string, RiskRate>();
                var currentBatch = new List<RiskRate>();

                // 라인 처리 및 중복 제거
                foreach (var line in lines)
                {
                    try
                    {
                        var parts = line.Split('\t');
                        if (parts.Length >= 6)
                        {
                            var riskRate = new RiskRate
                            {
                                COMPANY_NAME = parts[0],
                                RISK_RATE_NAME = parts[1].Replace(" ", "").Replace("\n", "").Replace("\r", "").Replace("\"", ""),
                                APPROVAL_DATE = parts[2],
                                FACTORS = parts[3],
                                RATES = parts[4],
                                COMMENT = parts[5],
                                CREATED_DATE = DateTimeOffset.UtcNow,
                                CREATED_BY = createdBy
                            };

                            // 복합키 생성
                            var key = $"{riskRate.COMPANY_NAME}|{riskRate.RISK_RATE_NAME}|{riskRate.APPROVAL_DATE}|{riskRate.FACTORS}";
                            uniqueRiskRates[key] = riskRate; // 중복된 키는 마지막 값으로 덮어씀
                        }
                        else
                        {
                            failedRegistrations.Add(new FailedRegistration
                            {
                                Line = line,
                                Reason = "데이터 형식 오류"
                            });
                            remainingLines.Remove(line);
                        }
                    }
                    catch (Exception ex)
                    {
                        failedRegistrations.Add(new FailedRegistration
                        {
                            Line = line,
                            Reason = $"처리 오류: {ex.Message}"
                        });
                        remainingLines.Remove(line);
                    }
                }

                // 중복 제거된 데이터를 배치로 처리
                var deduplicatedRiskRates = uniqueRiskRates.Values.ToList();
                var successCount = 0;

                for (int i = 0; i < deduplicatedRiskRates.Count; i += 1000)
                {
                    var batch = deduplicatedRiskRates.Skip(i).Take(1000).ToList();
                    try
                    {
                        var result = await _apiClient.CreateRiskRatesBatchAsync(batch);
                        successCount += result.SuccessCount;

                        // 실패한 항목 처리
                        foreach (var failure in result.Failures)
                        {
                            failedRegistrations.Add(new FailedRegistration
                            {
                                Line = FormatRiskRateToLine(failure),
                                Reason = "등록 실패"
                            });
                        }

                        // 진행상황 업데이트
                        UpdateProgress(successCount, totalCount, remainingLines, failedRegistrations);
                    }
                    catch (Exception ex)
                    {
                        foreach (var item in batch)
                        {
                            failedRegistrations.Add(new FailedRegistration
                            {
                                Line = FormatRiskRateToLine(item),
                                Reason = $"처리 오류: {ex.Message}"
                            });
                        }
                    }
                }

                // 최종 상태 업데이트
                var finalStatus = $"처리 완료: {successCount}/{totalCount} 건 성공 (중복 제외)";
                if (failedRegistrations.Any())
                {
                    finalStatus += $", {failedRegistrations.Count}건 실패";
                    var failedLines = failedRegistrations
                        .Select(f => $"{f.Line}\t[{f.Reason}]")
                        .ToList();
                    RegistrationTextBox.Text = string.Join(Environment.NewLine, failedLines);
                }
                else
                {
                    RegistrationTextBox.Clear();
                }

                RegistrationStatusText.Text = finalStatus;
            }
            catch (Exception ex)
            {
                RegistrationStatusText.Text = "처리 중 오류가 발생했습니다.";
                MessageBox.Show($"예상치 못한 오류가 발생했습니다. {ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                RegisterButton.IsEnabled = true;
                ClearButton.IsEnabled = true;
            }
        }

        private void UpdateProgress(int successCount, int totalCount, List<string> remainingLines, List<FailedRegistration> failedRegistrations)
        {
            // 진행 상황 텍스트 업데이트 (성공 건수와 실패 건수 모두 표시)
            RegistrationStatusText.Text = $"처리중: {successCount}/{totalCount} 건 완료, {failedRegistrations.Count}건 실패...";

            // 텍스트박스 업데이트
            var currentContent = new List<string>();

            // 아직 처리되지 않은 라인들
            lock (remainingLines)
            {
                currentContent.AddRange(remainingLines);
            }

            // 실패한 라인들 (실패 사유와 함께)
            lock (failedRegistrations)
            {
                currentContent.AddRange(failedRegistrations.Select(f => $"{f.Line}\t[{f.Reason}]"));
            }

            // 텍스트박스 업데이트
            RegistrationTextBox.Text = string.Join(Environment.NewLine, currentContent);
        }

        private string FormatRiskRateToLine(RiskRate riskRate)
        {
            return string.Join("\t",
                riskRate.COMPANY_NAME,
                riskRate.RISK_RATE_NAME,
                riskRate.APPROVAL_DATE,
                riskRate.FACTORS,
                riskRate.RATES,
                riskRate.COMMENT ?? "");
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            RegistrationTextBox.Clear();
            RegistrationStatusText.Text = string.Empty;
        }

        private async void GuideButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 실행파일 경로에서 가이드 파일 경로 생성
                string exePath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                string guidePath = Path.Combine(exePath, "risk_rates_postguide.xlsx");

                // 임시 폴더에 파일 복사
                string tempPath = Path.Combine(Path.GetTempPath(), "risk_rates_postguide.xlsx");

                // 파일 복사
                if (File.Exists(guidePath))
                {
                    using (var sourceStream = File.OpenRead(guidePath))
                    using (var fileStream = File.Create(tempPath))
                    {
                        await sourceStream.CopyToAsync(fileStream);
                    }

                    // 파일 실행
                    var processStartInfo = new ProcessStartInfo
                    {
                        FileName = tempPath,
                        UseShellExecute = true
                    };
                    Process.Start(processStartInfo);
                }
                else
                {
                    MessageBox.Show("가이드 파일을 찾을 수 없습니다.\n파일 경로: " + guidePath,
                        "파일 없음", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"가이드 파일을 여는 중 오류가 발생했습니다.\n{ex.Message}",
                    "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // 분석 탭 관련 메서드

        private void InitializeAnalysisTabContextMenu()
        {
            // DetailResultsGrid에 대한 ContextMenu 생성
            DetailResultsGrid.ContextMenu = new ContextMenu();

            // "분석 차트에 추가" 메뉴 아이템 생성
            var addToChartMenuItem = new MenuItem
            {
                Header = "분석 차트에 추가",
                Icon = new System.Windows.Controls.Image
                {
                    Source = new System.Windows.Media.Imaging.BitmapImage(
                        new Uri("/ActuLight;component/Resources/chart-add.png", UriKind.Relative))
                }
            };
            addToChartMenuItem.Click += AddToChart_Click;

            // ContextMenu에 메뉴 아이템 추가
            DetailResultsGrid.ContextMenu.Items.Add(addToChartMenuItem);

            // 컨텍스트 메뉴가 열릴 때 선택된 항목 확인
            DetailResultsGrid.ContextMenu.Opened += (s, e) =>
            {
                var selectedItems = DetailResultsGrid.SelectedItems;
                addToChartMenuItem.IsEnabled = selectedItems != null && selectedItems.Count > 0;

                // 선택된 항목 수에 따라 메뉴 텍스트 변경
                if (selectedItems != null && selectedItems.Count > 1)
                {
                    addToChartMenuItem.Header = $"선택한 {selectedItems.Count}개 항목을 분석 차트에 추가";
                }
                else
                {
                    addToChartMenuItem.Header = "분석 차트에 추가";
                }
            };
        }

        private void AddToChart_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = DetailResultsGrid.SelectedItems
                .Cast<RiskRateDetailViewModel>()
                .ToList();

            if (!selectedItems.Any())
                return;

            // 분석 탭으로 전환
            MainTabControl.SelectedItem = AnalysisTab;

            // 선택된 모든 항목을 차트에 추가
            foreach (var selectedRiskRate in selectedItems)
            {
                var result = _chartManager.AddRiskRate(selectedRiskRate);
                if (!result.Success)
                {
                    // 실패한 항목이 있을 경우, 어떤 항목이 실패했는지 표시
                    MessageBox.Show(
                        $"'{selectedRiskRate.RISK_RATE_NAME}' 추가 실패: {result.Message}",
                        "알림",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
            }
        }

        private void ApplyRangeButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (double.TryParse(MinXTextBox.Text, out double minX) &&
                    double.TryParse(MaxXTextBox.Text, out double maxX))
                {
                    if (minX >= maxX)
                    {
                        MessageBox.Show("MaxX 값은 MinX 값보다 커야 합니다.", "범위 오류",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    if (minX < 0)
                    {
                        MessageBox.Show("MinX 값은 0 이상이어야 합니다.", "범위 오류",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    if (maxX > 240)
                    {
                        MessageBox.Show("MaxX 값은 240 이하여야 합니다.", "범위 오류",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    _chartManager.UpdateXAxisRange(minX, maxX);
                    _chartManager.UpdateYAxisRange(minX, maxX);
                }
                else
                {
                    MessageBox.Show("올바른 숫자 형식을 입력해주세요.", "입력 오류",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"범위 적용 중 오류가 발생했습니다: {ex.Message}", "오류",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ClearChartButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _chartManager.ClearChart();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"차트 초기화 중 오류가 발생했습니다: {ex.Message}", "오류",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            _apiClient.Dispose();
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);

            // Ctrl+A로 전체 선택 지원
            if (e.Key == Key.A && Keyboard.Modifiers == ModifierKeys.Control)
            {
                if (DetailResultsGrid.IsFocused || DetailResultsGrid.IsKeyboardFocusWithin)
                {
                    DetailResultsGrid.SelectAll();
                    e.Handled = true;
                }
            }
        }
    }

    public class PaginatedResult<T>
    {
        [JsonProperty("items")]
        public List<T> Items { get; set; }

        [JsonProperty("totalCount")]
        public int TotalCount { get; set; }

        [JsonProperty("pageNumber")]
        public int PageNumber { get; set; }

        [JsonProperty("pageSize")]
        public int PageSize { get; set; }

        [JsonProperty("totalPages")]
        public int TotalPages { get; set; }

        [JsonProperty("hasNextPage")]
        public bool HasNextPage { get; set; }

        [JsonProperty("hasPreviousPage")]
        public bool HasPreviousPage { get; set; }
    }

    public class ApiClient : IDisposable
    {
        private readonly HttpClient _httpClient;
        private const string BaseUrl = "https://192.168.0.150:7186/";

        public ApiClient()
        {
            var handler = new HttpClientHandler
            {
                ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true
            };

            _httpClient = new HttpClient(handler)
            {
                BaseAddress = new Uri(BaseUrl)
            };

            _httpClient.DefaultRequestHeaders.Accept.Clear();
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        }

        // 회사명 목록 조회
        public async Task<List<string>> GetCompanyNamesAsync()
        {
            var response = await _httpClient.GetAsync("RiskRates/companies");
            await EnsureSuccessStatusCodeWithContentAsync(response);

            var content = await response.Content.ReadAsStringAsync();
            return JsonConvert.DeserializeObject<List<string>>(content);
        }

        // 검색 기능
        public async Task<PaginatedResult<RiskRate>> SearchRiskRatesAsync(string searchTerm, int pageNumber = 1, int pageSize = 20, string company = null)
        {
            // searchTerm이 비어있으면 "empty" 키워드 사용
            var searchParam = string.IsNullOrWhiteSpace(searchTerm) ? "empty" : Uri.EscapeDataString(searchTerm);
            var url = $"RiskRates/search?searchTerm={searchParam}&pageNumber={pageNumber}&pageSize={pageSize}";

            if (!string.IsNullOrEmpty(company))
            {
                url += $"&company={Uri.EscapeDataString(company)}";
            }

            var response = await _httpClient.GetAsync(url);
            await EnsureSuccessStatusCodeWithContentAsync(response);

            var content = await response.Content.ReadAsStringAsync();
            return JsonConvert.DeserializeObject<PaginatedResult<RiskRate>>(content);
        }

        // 단일 항목 조회
        public async Task<RiskRate> GetRiskRateAsync(string companyName, string rateName, string approvalDate, string factors)
        {
            var response = await _httpClient.GetAsync(
                $"RiskRates/{Uri.EscapeDataString(companyName)}/" +
                $"{Uri.EscapeDataString(rateName)}/" +
                $"{Uri.EscapeDataString(approvalDate)}/" +
                $"{Uri.EscapeDataString(factors)}");

            await EnsureSuccessStatusCodeWithContentAsync(response);

            var content = await response.Content.ReadAsStringAsync();
            return JsonConvert.DeserializeObject<RiskRate>(content);
        }

        // 자동완성
        public async Task<List<string>> GetRiskRateNamesAsync(string searchTerm, string company = null)
        {
            try
            {
                var url = $"RiskRates/riskrates?searchTerm={Uri.EscapeDataString(searchTerm)}";
                if (!string.IsNullOrEmpty(company))
                {
                    url += $"&company={Uri.EscapeDataString(company)}";
                }

                var response = await _httpClient.GetAsync(url);
                await EnsureSuccessStatusCodeWithContentAsync(response);

                var content = await response.Content.ReadAsStringAsync();
                return JsonConvert.DeserializeObject<List<string>>(content);
            }
            catch (Exception ex)
            {
                throw new ApiException("Failed to get risk rate names", System.Net.HttpStatusCode.InternalServerError, ex.Message);
            }
        }

        // 이름 변경
        public async Task RenameRiskRateAsync(string companyName, string oldName, string approvalDate, string factors, string newName)
        {
            var url = $"RiskRates/rename?companyName={Uri.EscapeDataString(companyName)}" +
                      $"&oldName={Uri.EscapeDataString(oldName)}" +
                      $"&approvalDate={Uri.EscapeDataString(approvalDate)}" +
                      $"&factors={Uri.EscapeDataString(factors)}" +
                      $"&newName={Uri.EscapeDataString(newName)}";

            var response = await _httpClient.PutAsync(url, null);
            await EnsureSuccessStatusCodeWithContentAsync(response);
        }

        // 등록
        public async Task<RiskRate> CreateRiskRateAsync(RiskRate riskRate) 
        {
            var json = JsonConvert.SerializeObject(riskRate);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _httpClient.PostAsync("RiskRates", content);
            await EnsureSuccessStatusCodeWithContentAsync(response);

            var responseContent = await response.Content.ReadAsStringAsync();
            return JsonConvert.DeserializeObject<RiskRate>(responseContent);
        }

        public async Task<BatchResult> CreateRiskRatesBatchAsync(List<RiskRate> riskRates)
        {
            var json = JsonConvert.SerializeObject(riskRates);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _httpClient.PostAsync("RiskRates/batch", content);
            await EnsureSuccessStatusCodeWithContentAsync(response);

            var responseContent = await response.Content.ReadAsStringAsync();
            return JsonConvert.DeserializeObject<BatchResult>(responseContent);
        }

        public async Task<BatchResult> UpdateRiskRatesBatchAsync(List<RiskRate> riskRates)
        {
            var json = JsonConvert.SerializeObject(riskRates);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _httpClient.PutAsync("RiskRates/batch", content);
            await EnsureSuccessStatusCodeWithContentAsync(response);

            var responseContent = await response.Content.ReadAsStringAsync();
            return JsonConvert.DeserializeObject<BatchResult>(responseContent);
        }

        // 일괄 등록
        public async Task<List<RiskRate>> CreateRiskRatesAsync(List<RiskRate> riskRates)
            {
                var createdRates = new List<RiskRate>();
                foreach (var rate in riskRates)
                {
                    var createdRate = await CreateRiskRateAsync(rate);
                    createdRates.Add(createdRate);
                }
                return createdRates;
            }

        // 수정
        public async Task UpdateRiskRateAsync(RiskRate riskRate)
        {
            var json = JsonConvert.SerializeObject(riskRate);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _httpClient.PutAsync("RiskRates", content);
            await EnsureSuccessStatusCodeWithContentAsync(response);
        }

        // 삭제
        public async Task DeleteRiskRateAsync(string companyName, string rateName, string approvalDate, string factors)
        {
            var response = await _httpClient.DeleteAsync(
                $"RiskRates/{Uri.EscapeDataString(companyName)}/" +
                $"{Uri.EscapeDataString(rateName)}/" +
                $"{Uri.EscapeDataString(approvalDate)}/" +
                $"{Uri.EscapeDataString(factors)}");

            await EnsureSuccessStatusCodeWithContentAsync(response);
        }

        // 회사별 조회
        public async Task<List<RiskRate>> GetRiskRatesByCompanyAsync(string companyName)
        {
            var response = await _httpClient.GetAsync($"RiskRates/company/{Uri.EscapeDataString(companyName)}");
            await EnsureSuccessStatusCodeWithContentAsync(response);

            var content = await response.Content.ReadAsStringAsync();
            return JsonConvert.DeserializeObject<List<RiskRate>>(content);
        }

        // API 오류 처리 헬퍼 메서드
        private async Task EnsureSuccessStatusCodeWithContentAsync(HttpResponseMessage response)
        {
            if (!response.IsSuccessStatusCode)
            {
                var content = await response.Content.ReadAsStringAsync();
                throw new ApiException(
                    $"API request failed with status code: {response.StatusCode}",
                    response.StatusCode,
                    content
                );
            }
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }

        public class BatchResult
        {
            public int SuccessCount { get; set; } = 0;
            public List<RiskRate> NotFound { get; set; } = new List<RiskRate>();
            public List<RiskRate> Failures { get; set; } = new List<RiskRate>();
        }
    }

    public class ApiException : Exception
    {
        public System.Net.HttpStatusCode StatusCode { get; }
        public string ResponseContent { get; }

        public ApiException(string message, System.Net.HttpStatusCode statusCode, string responseContent)
            : base(message)
        {
            StatusCode = statusCode;
            ResponseContent = responseContent;
        }
    }

    public class RiskRate
    {
        // Primary Key 
        public string COMPANY_NAME { get; set; }
        public string RISK_RATE_NAME { get; set; }
        public string APPROVAL_DATE { get; set; }
        public string FACTORS { get; set; }

        // Normal Fields
        public string RATES { get; set; }
        public string COMMENT { get; set; }

        // Audit Fields
        public DateTimeOffset CREATED_DATE { get; set; }
        public string CREATED_BY { get; set; }
    }

    public class RiskRateSummary
    {
        public string 회사명 { get; set; }
        public string 위험률명 { get; set; }
        public string 적용일자 { get; set; }
        public string 기간구조 { get; set; }
        public string 팩터구성 { get; set; }
        public string 등록일시 { get; set; }
        public string 등록자 { get; set; }
        public List<RiskRate> Details { get; set; } = new List<RiskRate>();

        public static RiskRateSummary FromRiskRate(RiskRate rate)
        {
            string[] factors = rate.FACTORS.Split('|');
            var factorString = string.Join(" ",
                Enumerable.Range(1, factors.Length - 1)
                    .Where(i => !string.IsNullOrWhiteSpace(factors[i]))
                    .Select(i => $"F{i}"));

            return new RiskRateSummary
            {
                회사명 = rate.COMPANY_NAME,
                위험률명 = rate.RISK_RATE_NAME,
                적용일자 = rate.APPROVAL_DATE,
                기간구조 = factors[0],
                팩터구성 = factorString,
                등록일시 = rate.CREATED_DATE.LocalDateTime.ToString("yyyy-MM-dd HH:mm"),
                등록자 = rate.CREATED_BY.ToString(),
                Details = new List<RiskRate> { rate }
            };
        }

        public string Key => string.Join("|", 회사명, 위험률명, 적용일자, 기간구조, 팩터구성);
    }

    public class RiskRateDetailViewModel
    {
        public string COMPANY_NAME { get; set; }
        public string RISK_RATE_NAME { get; set; }
        public string APPROVAL_DATE { get; set; }
        public string FACTORS { get; set; }
        public string RATES { get; set; }
        public string COMMENT { get; set; }
        public string 기간구분 { get; set; }
        public string F1 { get; set; }
        public string F2 { get; set; }
        public string F3 { get; set; }
        public string F4 { get; set; }
        public string F5 { get; set; }
        public string F6 { get; set; }
        public string F7 { get; set; }
        public string F8 { get; set; }
        public string F9 { get; set; }
        public string[] RateValues { get; set; } = new string[241];
    }

    public class FailedRegistration
    {
        public string Line { get; set; }
        public string Reason { get; set; }
    }

    // 팩터 값들을 비교하기 위한 IComparer 구현
    public class FactorComparer : IComparer<List<int>>
    {
        public int Compare(List<int> x, List<int> y)
        {
            if (x == null || y == null)
                return 0;

            for (int i = 0; i < Math.Min(x.Count, y.Count); i++)
            {
                int comparison = x[i].CompareTo(y[i]);
                if (comparison != 0)
                    return comparison;
            }

            return x.Count.CompareTo(y.Count);
        }
    }

    public class RiskRateChartManager
    {
        private readonly CartesianChart _chart;
        private readonly Dictionary<string, LineSeries> _chartSeries;
        private readonly List<Color> _colorPalette;

        public RiskRateChartManager(CartesianChart chart)
        {
            _chart = chart ?? throw new ArgumentNullException(nameof(chart));
            _chartSeries = new Dictionary<string, LineSeries>();
            _colorPalette = InitializeColorPalette();

            InitializeChartConfiguration();
        }

        private List<Color> InitializeColorPalette() => new List<Color>
        {
            Color.FromRgb(0, 150, 255),    // 밝은 파랑
            Color.FromRgb(255, 82, 82),    // 산호색
            Color.FromRgb(0, 219, 158),    // 민트
            Color.FromRgb(255, 159, 64),   // 주황
            Color.FromRgb(156, 39, 176),   // 보라
            Color.FromRgb(255, 193, 7),    // 노랑
            Color.FromRgb(233, 30, 99),    // 분홍
            Color.FromRgb(0, 230, 118),    // 초록
            Color.FromRgb(103, 58, 183),   // 인디고
            Color.FromRgb(1, 192, 200)     // 청록
        };

        private void InitializeChartConfiguration()
        {
            // 기본 차트 설정
            ConfigureChartBase();

            // 축 설정
            ConfigureAxes();

            // 시각화 설정
            ConfigureVisualization();

            // 성능 최적화 설정
            ConfigurePerformanceOptimizations();
        }

        private void ConfigureChartBase()
        {
            _chart.Series = new SeriesCollection();
            _chart.DisableAnimations = true;
            _chart.AnimationsSpeed = TimeSpan.Zero;
            _chart.Zoom = ZoomingOptions.None;
            _chart.Pan = PanningOptions.None;
            _chart.Hoverable = true; // 호버 효과 비활성화
            _chart.Background = new SolidColorBrush(Color.FromRgb(28, 28, 30));
            _chart.Foreground = new SolidColorBrush(Colors.WhiteSmoke);
            _chart.LegendLocation = LegendLocation.Right;

        }

        private void ConfigureAxes()
        {
            // X축 설정
            _chart.AxisX = new AxesCollection
            {
                new Axis
                {
                    Title = "기간",
                    FontSize = 11,
                    FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
                    Foreground = new SolidColorBrush(Colors.WhiteSmoke),
                    FontWeight = FontWeights.SemiBold,
                    MinValue = 0,
                    MaxValue = 130,
                    MinRange = 5,
                    MaxRange = 130,
                    Separator = new LiveCharts.Wpf.Separator
                    {
                        Stroke = new SolidColorBrush(Color.FromRgb(64, 64, 64)),
                        StrokeThickness = 0.5,
                        Step = 10
                    },
                    LabelFormatter = value => ((int)value % 10 == 0) ? value.ToString() : ""
                }
            };

            // Y축 설정
            _chart.AxisY = new AxesCollection
            {
                new Axis
                {
                    Title = "위험률",
                    FontSize = 11,
                    FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
                    Foreground = new SolidColorBrush(Colors.WhiteSmoke),
                    FontWeight = FontWeights.SemiBold,
                    MinValue = 0,  // Y축 최소값을 0으로 설정
                    LabelFormatter = value => value.ToString("F4"),
                    Separator = new LiveCharts.Wpf.Separator
                    {
                        Stroke = new SolidColorBrush(Color.FromRgb(64, 64, 64)),
                        StrokeThickness = 0.5
                    }
                }
            };
        }

        private void ConfigureVisualization()
        {
            _chart.DataTooltip = new DefaultTooltip
            {
                SelectionMode = TooltipSelectionMode.OnlySender,
                ShowTitle = false,
                ShowSeries = true,
                CornerRadius = new CornerRadius(3),
                Background = new SolidColorBrush(Color.FromRgb(45, 45, 48)),
                Foreground = new SolidColorBrush(Colors.WhiteSmoke),
                FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
                FontSize = 11
            };
        }

        private void ConfigurePerformanceOptimizations()
        {
            RenderOptions.SetCachingHint(_chart, CachingHint.Cache);
            ScrollViewer.SetCanContentScroll(_chart, true);
        }

        public AddChartResult AddRiskRate(RiskRateDetailViewModel riskRate)
        {
            string seriesKey = GetSeriesKey(riskRate);

            if (_chartSeries.ContainsKey(seriesKey))
            {
                return new AddChartResult
                {
                    Success = false,
                    Message = "이미 추가된 위험률입니다."
                };
            }

            var values = ParseRateValues(riskRate.RateValues);
            if (!values.Any())
            {
                return new AddChartResult
                {
                    Success = false,
                    Message = "유효한 위험률 데이터가 없습니다."
                };
            }

            var series = CreateSeries(riskRate, values);
            _chart.Series.Add(series);
            _chartSeries[seriesKey] = series;

            // 시리즈 추가 후 현재 X축 범위에 맞춰 Y축 업데이트
            if (_chart.AxisX.Any())
            {
                var xAxis = _chart.AxisX[0];
                UpdateYAxisRange(xAxis.MinValue, xAxis.MaxValue);
            }

            return new AddChartResult { Success = true };
        }

        public void RemoveRiskRate(string seriesKey)
        {
            if (_chartSeries.TryGetValue(seriesKey, out var series))
            {
                _chart.Series.Remove(series);
                _chartSeries.Remove(seriesKey);

                // 시리즈 제거 후 현재 X축 범위에 맞춰 Y축 업데이트
                if (_chart.AxisX.Any())
                {
                    var xAxis = _chart.AxisX[0];
                    UpdateYAxisRange(xAxis.MinValue, xAxis.MaxValue);
                }
            }
        }

        public void UpdateXAxisRange(double minX, double maxX)
        {
            if (_chart.AxisX.Any())
            {
                var axis = _chart.AxisX[0];
                axis.MinValue = minX;
                axis.MaxValue = maxX;
                axis.MinRange = Math.Max(5, (maxX - minX) * 0.1); // 최소 범위는 5 또는 전체 범위의 10%
                axis.MaxRange = maxX - minX;
            }
        }

        public void UpdateYAxisRange(double minX, double maxX)
        {
            if (!_chart.Series.Any()) return;

            var visibleValues = new List<double>();
            var maxXValue = Math.Min(130, maxX); // 최대값을 130으로 제한

            foreach (var series in _chart.Series)
            {
                var values = ((ChartValues<double>)series.Values)
                    .Select((value, index) => new { Value = value, Index = index })
                    .Where(x => x.Index >= minX && x.Index <= maxXValue)
                    .Select(x => x.Value);

                visibleValues.AddRange(values);
            }

            if (visibleValues.Any())
            {
                var minY = visibleValues.Min();
                var maxY = visibleValues.Max();
                var padding = (maxY - minY) * 0.1;

                if (_chart.AxisY.Any())
                {
                    var yAxis = _chart.AxisY[0];
                    yAxis.MinValue = 0.0;
                    yAxis.MaxValue = maxY + padding;
                }
            }
        }

        private LineSeries CreateSeries(RiskRateDetailViewModel riskRate, ChartValues<double> values)
        {
            var color = _colorPalette[_chartSeries.Count % _colorPalette.Count];

            var series = new LineSeries
            {
                Title = FormatSeriesTitle(riskRate),
                Values = values,
                Stroke = new SolidColorBrush(color),
                Fill = Brushes.Transparent,
                PointGeometry = DefaultGeometries.Circle,
                PointGeometrySize = 5,
                LineSmoothness = 0,
                StrokeThickness = 1.5,
                IsHitTestVisible = true,
                DataLabels = false,
                LabelPoint = point => $"x: {point.X:F0}, y: {point.Y}"
            };

            return series;
        }

        public void HighlightSeries(LineSeries selectedSeries)
        {
            foreach (var series in _chart.Series.OfType<LineSeries>())
            {
                var isSelected = series == selectedSeries;
                var color = ((SolidColorBrush)series.Stroke).Color;

                series.Stroke = new SolidColorBrush(color)
                {
                    Opacity = isSelected ? 1 : 0.3
                };

                series.Fill = new SolidColorBrush(Color.FromArgb(
                    (byte)(isSelected ? 30 : 10),
                    color.R,
                    color.G,
                    color.B
                ));

                series.StrokeThickness = isSelected ? 3 : 1;
                series.PointGeometrySize = isSelected ? 8 : 4;
            }
        }

        public void ResetHighlight()
        {
            foreach (var series in _chart.Series.OfType<LineSeries>())
            {
                var color = ((SolidColorBrush)series.Stroke).Color;
                series.Stroke = new SolidColorBrush(color) { Opacity = 1 };
                series.Fill = new SolidColorBrush(Color.FromArgb(20, color.R, color.G, color.B));
                series.StrokeThickness = 2;
                series.PointGeometrySize = 6;
            }
        }

        private string GetSeriesKey(RiskRateDetailViewModel riskRate)
        {
            return $"{riskRate.COMPANY_NAME}_{riskRate.RISK_RATE_NAME}_{riskRate.APPROVAL_DATE}_{riskRate.FACTORS}";
        }

        private ChartValues<double> ParseRateValues(string[] rateValues)
        {
            var values = new ChartValues<double>();
            // 0~130까지만 파싱
            for (int i = 0; i < Math.Min(131, rateValues.Length); i++)
            {
                if (double.TryParse(rateValues[i], out double value))
                {
                    values.Add(value);
                }
                else
                {
                    values.Add(0);
                }
            }
            return values;
        }

        private string FormatSeriesTitle(RiskRateDetailViewModel riskRate)
        {
            var firstLine = $"{riskRate.COMPANY_NAME} - {riskRate.RISK_RATE_NAME}";
            var secondLine = $"({riskRate.APPROVAL_DATE}, {FormatFactors(riskRate.FACTORS)})";
            return $"{firstLine}\n{secondLine}";
        }

        private string FormatFactors(string factors)
        {
            var factorParts = factors.Split('|');
            var validFactors = factorParts
                .Select((value, index) => new { Value = value, Index = index })
                .Where(x => !string.IsNullOrWhiteSpace(x.Value))
                .Select(x => x.Index == 0 ? x.Value : $"F{x.Index}:{x.Value}")
                .ToList();

            return string.Join(", ", validFactors);
        }

        public void ClearChart()
        {
            _chart.Series.Clear();
            _chartSeries.Clear();

            // Y축 범위 초기화
            _chart.AxisY[0].MinValue = 0;
            _chart.AxisY[0].MaxValue = 1;
        }
    }

    public class AddChartResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }
    }
}

