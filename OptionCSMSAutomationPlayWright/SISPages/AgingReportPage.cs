using Microsoft.Playwright;
using OfficeOpenXml;
using System.Globalization;
using System.Text.RegularExpressions;
using IPage = Microsoft.Playwright.IPage;
using Path = System.IO.Path;

namespace OptionCSMSAutomationPlayWright.SISPages
{
    /// <summary>
    /// Acutis dynamic report (Id=106): shows every record, sorts by Aging Days (largest first),
    /// highlights Aging Days above a limit, takes a full-page screenshot and inserts it into the
    /// "Aging Report" sheet of today's audit workbook.
    /// </summary>
    public class AgingReportPage
    {
        private const string AgingSheetName = "Aging Report";
        private const int MaxScreenshotChunkHeight = 8000;   // Chromium cannot reliably capture much more per shot
        private const int ExcelRowHeightPixels = 20;         // default Excel row = 15pt = 20px
        private const string AgingColumnPattern = "aging";

        private readonly IPage _page;
        private readonly string directoryPath = "D:\\File";
        private readonly string auditWorkbookPath;
        private int _highlighted;
        private int _totalRecords;
        private int _nextExcelRow;                 // first free 1-based row of the "Aging Report" sheet
        private DateTime _serviceFeeStart;
        private DateTime _serviceFeeEnd;
        private string _serviceFeeText = string.Empty;
        private string _autoWithdrawalText = string.Empty;
        private System.Drawing.Color _autoWithdrawalColor = System.Drawing.Color.DarkBlue;

        public AgingReportPage(IPage page)
        {
            _page = page;
            Directory.CreateDirectory(directoryPath);
            string datePrefix = MMAuditPages.GetIndiaToday().ToString("MMddyyyy", CultureInfo.InvariantCulture);
            auditWorkbookPath = Path.Combine(directoryPath, $"{datePrefix}- AuditSummary.xlsx");
        }

        public async Task OpenReportAsync(string url)
        {
            await _page.GotoAsync(url, new() { Timeout = 120000, WaitUntil = WaitUntilState.DOMContentLoaded });
            try { await _page.WaitForLoadStateAsync(LoadState.NetworkIdle, new() { Timeout = 60000 }); } catch (TimeoutException) { }
            await _page.Locator("th").Filter(new() { HasTextRegex = new Regex(AgingColumnPattern, RegexOptions.IgnoreCase) })
                .First.WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = 120000 });
            await WaitForRowsToSettleAsync();
            Console.WriteLine($"Dynamic report opened: {url} ({await CountBodyRowsAsync()} row(s) visible).");
        }

        public async Task ShowAllRecordsAsync()
        {
            // The "records" page-size select (DataTables "<id>_length") may be hidden behind a styled widget,
            // so choose the "All" option directly and raise the change event DataTables listens to.
            bool chosen = await _page.EvaluateAsync<bool>(@"() => {
                const sel = [...document.querySelectorAll('select')].find(s => [...s.options].some(o => /^\s*all\s*$/i.test(o.text)));
                if (!sel) return false;
                const opt = [...sel.options].find(o => /^\s*all\s*$/i.test(o.text));
                sel.value = opt.value;
                if (window.jQuery) { window.jQuery(sel).trigger('change'); }
                else { sel.dispatchEvent(new Event('change', { bubbles: true })); }
                return true;
            }");
            if (!chosen)
            {
                string found = await _page.EvaluateAsync<string>(
                    "[...document.querySelectorAll('select')].map(s => (s.name||s.id||s.className) + ': ' + [...s.options].map(o => o.value + '=' + o.text).join(',')).join(' | ')");
                throw new InvalidOperationException("No paging drop-down with an 'All' option was found on the report. Drop-downs on page: " + found);
            }
            await WaitForRowsToSettleAsync(120000);
            _totalRecords = await CountBodyRowsAsync();
            Console.WriteLine($"Paging set to All: {_totalRecords} record(s) displayed.");
        }

        public async Task SortByAgingDaysDescendingAsync()
        {
            var header = _page.Locator("th").Filter(new() { HasTextRegex = new Regex(AgingColumnPattern, RegexOptions.IgnoreCase) }).First;
            for (int attempt = 1; attempt <= 4; attempt++)
            {
                var values = await ReadAgingValuesAsync();
                if (values.Count > 1 && IsDescending(values))
                {
                    Console.WriteLine($"Sorted by Aging Days, largest first (max {values.First()}, min {values.Last()}).");
                    return;
                }
                await header.ClickAsync();
                await WaitForRowsToSettleAsync(60000);
            }
            var finalValues = await ReadAgingValuesAsync();
            if (finalValues.Count > 1 && !IsDescending(finalValues))
                throw new InvalidOperationException("Could not sort the Aging Days column in descending order.");
        }

        public async Task HighlightAgingDaysAboveAsync(int limit)
        {
            // Also lets the full table render in the screenshot: remove fixed-height scroll containers around it.
            _highlighted = await _page.EvaluateAsync<int>(@"(args) => {
                const limit = args.limit, pat = args.pat;
                const head = [...document.querySelectorAll('th')].find(h => new RegExp(pat, 'i').test(h.textContent));
                if (!head) return -1;
                const idx = head.cellIndex;
                const tables = [...document.querySelectorAll('table')].filter(t => t.querySelectorAll('tbody tr td').length > 0);
                const table = tables.find(t => t.contains(head)) || tables[0];
                let count = 0;
                table.querySelectorAll('tbody tr').forEach(tr => {
                    const cell = tr.cells[idx];
                    if (!cell) return;
                    const n = parseInt((cell.textContent || '').replace(/[^0-9-]/g, ''), 10);
                    if (!isNaN(n) && n > limit) {
                        cell.style.setProperty('background-color', '#ffff00', 'important');
                        cell.style.setProperty('color', '#c00000', 'important');
                        cell.style.setProperty('font-weight', 'bold', 'important');
                        count++;
                    }
                });
                for (let el = table.parentElement; el && el !== document.documentElement; el = el.parentElement) {
                    const cs = getComputedStyle(el);
                    if (/(auto|scroll|hidden)/.test(cs.overflowY + cs.overflow) || cs.maxHeight !== 'none' || el.style.height) {
                        el.style.setProperty('max-height', 'none', 'important');
                        el.style.setProperty('height', 'auto', 'important');
                        el.style.setProperty('overflow', 'visible', 'important');
                    }
                }
                return count;
            }", new { limit, pat = AgingColumnPattern });

            if (_highlighted < 0)
                throw new InvalidOperationException("The Aging Days column was not found, nothing was highlighted.");
            Console.WriteLine($"Highlighted {_highlighted} record(s) with Aging Days greater than {limit}.");
        }

        public async Task InsertScreenshotInNextSheetAsync(int limit)
        {
            var files = await CaptureGridAsync("Aging Report");
            InsertPicturesInWorkbook(files, limit);
        }

        private void InsertPicturesInWorkbook(List<string> files, int limit)
        {
            try
            {
                using var package = new ExcelPackage(new FileInfo(auditWorkbookPath));
                var existing = package.Workbook.Worksheets.FirstOrDefault(w => w.Name == AgingSheetName);
                if (existing != null) package.Workbook.Worksheets.Delete(existing);
                var ws = package.Workbook.Worksheets.Add(AgingSheetName);
                ws.View.ShowGridLines = false;

                ws.Cells["A1"].Value = $"Aging Report (dynamic report Id=106) - {DateTime.Now:MM/dd/yyyy hh:mm tt}";
                ws.Cells["A1"].Style.Font.Bold = true;
                ws.Cells["A2"].Value = $"Total records: {_totalRecords}. Aging Days greater than {limit} are highlighted: {_highlighted}.";

                int nextRow = 3; // zero-based row index for the picture anchor
                for (int i = 0; i < files.Count; i++)
                {
                    var pic = ws.Drawings.AddPicture($"AgingReport{i + 1}", new FileInfo(files[i]));
                    pic.SetPosition(nextRow, 0, 0, 0);
                    nextRow += (int)Math.Ceiling(pic.Image.Bounds.Height / (double)ExcelRowHeightPixels) + 1;
                }
                _nextExcelRow = nextRow + 2; // first free 1-based row below the picture(s), after one blank row
                package.Save();
            }
            catch (Exception ex) when (ex is IOException || ex.ToString().Contains("being used by another process"))
            {
                throw new IOException($"Could not write the screenshot into '{auditWorkbookPath}'. Close the workbook if it is open in Excel and run again. {ex.Message}", ex);
            }
            DeleteScreenshotFiles(files); // the pictures are now stored inside the workbook
            Console.WriteLine($"Screenshot inserted in sheet '{AgingSheetName}' of {auditWorkbookPath}.");
        }

        // ------------------------------------------------------------------------------------------------
        // eCheck Service Fee Report
        // ------------------------------------------------------------------------------------------------

        public async Task OpenReportsMenuItemAsync(string linkId, string menuName)
        {
            // the menu markup exists twice (desktop and a hidden copy), so only look at the visible one
            var reportsMenu = _page.Locator("a[href='#']:visible").Filter(new() { HasTextRegex = new Regex(@"^\s*Reports\s*$") }).First;
            var item = _page.Locator($"a#{linkId}:visible").First;
            await reportsMenu.HoverAsync();
            try { await item.WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = 3000 }); }
            catch (TimeoutException)
            {
                await reportsMenu.ClickAsync(); // menu opens on click when hover does not show it
                await item.WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = 10000 });
            }
            await item.ClickAsync();
            try { await _page.WaitForLoadStateAsync(LoadState.NetworkIdle, new() { Timeout = 60000 }); } catch (TimeoutException) { }
            Console.WriteLine($"Opened Reports > {menuName}.");
        }

        public async Task OpenServiceFeeReportTabAsync()
        {
            var tab = _page.Locator("a.nav-link").Filter(new() { HasTextRegex = new Regex(@"^\s*Service Fee Report\s*$") }).First;
            await tab.ClickAsync();
            await _page.Locator("#txtStartDate").WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = 30000 });
            await _page.Locator("h2.pageinner_title").Filter(new() { HasTextRegex = new Regex("SERVICE FEE", RegexOptions.IgnoreCase) })
                .First.WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = 30000 });
            Console.WriteLine("Service Fee Report tab opened.");
        }

        /// <summary>Yesterday, or Friday to Sunday of the previous week when run on a Monday.</summary>
        public static (DateTime Start, DateTime End) GetServiceFeeDateRange(DateTime today)
            => today.DayOfWeek == DayOfWeek.Monday
                ? (today.AddDays(-3), today.AddDays(-1))
                : (today.AddDays(-1), today.AddDays(-1));

        public async Task FilterServiceFeeReportAsync()
        {
            (_serviceFeeStart, _serviceFeeEnd) = GetServiceFeeDateRange(DateTime.Today);
            await SetDatePickerAsync("#txtStartDate", _serviceFeeStart);
            await SetDatePickerAsync("#txtEndDate", _serviceFeeEnd);

            await _page.Locator("input[value='Run Report']").ClickAsync();
            await _page.Locator("text=/Showing \\d+ to \\d+ of/").First.WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = 180000 });
            await Task.Delay(2000);
            Console.WriteLine($"Service Fee Report run for {_serviceFeeStart:MM/dd/yyyy} to {_serviceFeeEnd:MM/dd/yyyy}.");
        }

        public async Task ResizePageToFitReportAsync()
        {
            int height = await _page.EvaluateAsync<int>("Math.max(document.documentElement.scrollHeight, document.body.scrollHeight)");
            await _page.SetViewportSizeAsync(1500, Math.Min(height + 20, 8000));
            await Task.Delay(700);
            Console.WriteLine($"Page resized to 1500x{Math.Min(height + 20, 8000)} to fit the whole report.");
        }

        public async Task HighlightServiceFeeTotalAsync()
        {
            // Yellow fill and a thick yellow border around the Total row (table footer).
            bool found = await _page.EvaluateAsync<bool>(@"() => {
                const rows = [...document.querySelectorAll('tfoot tr')].filter(r => /total/i.test(r.innerText));
                if (!rows.length) return false;
                rows.forEach(r => [...r.cells].forEach(td => {
                    td.style.setProperty('background-color', '#ffff00', 'important');
                    td.style.setProperty('border-top', '5px solid #ffd000', 'important');
                    td.style.setProperty('border-bottom', '5px solid #ffd000', 'important');
                    td.style.setProperty('font-weight', 'bold', 'important');
                }));
                const first = rows[0].cells[0], last = rows[0].cells[rows[0].cells.length - 1];
                first.style.setProperty('border-left', '5px solid #ffd000', 'important');
                last.style.setProperty('border-right', '5px solid #ffd000', 'important');
                return true;
            }");
            Console.WriteLine(found
                ? "Service Fee Total row highlighted (yellow, thick border)."
                : "No Total row on the Service Fee report (no data for the selected dates); nothing to highlight.");
        }

        public async Task DisplayServiceFeeTextAsync()
        {
            _serviceFeeText = _serviceFeeStart.Date == _serviceFeeEnd.Date
                ? $"Service Fee to OptionC on {_serviceFeeStart:MM/dd/yyyy}"
                : $"Service Fee to OptionC on {_serviceFeeStart:MM/dd/yyyy} to {_serviceFeeEnd:MM/dd/yyyy}";
            await ShowBannerAsync(_serviceFeeText, "#ff0000");
            Console.WriteLine($"Displayed text: {_serviceFeeText}");
        }

        public async Task AttachServiceFeeScreenshotAsync()
        {
            var files = await CaptureGridAsync("Service Fee Report");
            AppendSectionToSheet(_serviceFeeText, files, System.Drawing.Color.Red);
        }

        // ------------------------------------------------------------------------------------------------
        // Auto Withdrawals
        // ------------------------------------------------------------------------------------------------

        public async Task FilterAutoWithdrawalsForTodayAsync()
        {
            DateTime today = DateTime.Today;
            await _page.Locator("#StartDate").WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = 30000 });
            await SetDatePickerAsync("#StartDate", today);
            await SetDatePickerAsync("#EndDate", today);
            await _page.Locator("input[value='Filter']").ClickAsync();
            try { await _page.WaitForLoadStateAsync(LoadState.NetworkIdle, new() { Timeout = 60000 }); } catch (TimeoutException) { }
            await WaitForInfoTextToSettleAsync();

            int count = await ReadAutoWithdrawalCountAsync();
            if (count > 25)
            {
                // show every transaction so the screenshot holds the whole list
                await _page.EvaluateAsync(@"() => { const sel = [...document.querySelectorAll('select')].find(s => [...s.options].some(o => /^\s*all\s*$/i.test(o.text)));
                    if (!sel) return; sel.value = [...sel.options].find(o => /^\s*all\s*$/i.test(o.text)).value;
                    if (window.jQuery) { window.jQuery(sel).trigger('change'); } else { sel.dispatchEvent(new Event('change', { bubbles: true })); } }");
                await WaitForInfoTextToSettleAsync();
            }

            if (count > 0)
            {
                _autoWithdrawalText = $"Auto withdrawal transactions scheduled today ({today:MM/dd/yyyy}): {count}";
                _autoWithdrawalColor = System.Drawing.Color.DarkBlue;
            }
            else
            {
                _autoWithdrawalText = $"No auto withdrawal transaction scheduled today ({today:MM/dd/yyyy})";
                _autoWithdrawalColor = System.Drawing.Color.FromArgb(0xC0, 0x00, 0x00);
            }
            Console.WriteLine($"Auto withdrawals for {today:MM/dd/yyyy}: {count}.");
        }

        public async Task DisplayAutoWithdrawalTextAsync()
        {
            string css = $"#{_autoWithdrawalColor.R:X2}{_autoWithdrawalColor.G:X2}{_autoWithdrawalColor.B:X2}";
            await ShowBannerAsync(_autoWithdrawalText, css);
            Console.WriteLine($"Displayed text: {_autoWithdrawalText}");
        }

        public async Task AttachAutoWithdrawalScreenshotAsync()
        {
            var files = await CaptureGridAsync("Auto Withdrawals");
            AppendSectionToSheet(_autoWithdrawalText, files, _autoWithdrawalColor);
        }

        private async Task<int> ReadAutoWithdrawalCountAsync()
        {
            string info = await _page.Locator("text=/Showing \\d+ to \\d+ of/").First.InnerTextAsync();
            var m = Regex.Match(info, @"of\s+([\d,]+)\s+entries");
            int fromInfo = m.Success ? int.Parse(m.Groups[1].Value.Replace(",", ""), CultureInfo.InvariantCulture) : -1;
            if (fromInfo >= 0) return fromInfo;
            // fall back to the rows themselves (an empty grid has one "No data available" row)
            int rows = await _page.Locator("table tbody tr:has(td):not(:has(td.dataTables_empty))").CountAsync();
            return rows;
        }

        // ------------------------------------------------------------------------------------------------
        // Shared helpers
        // ------------------------------------------------------------------------------------------------

        // These date pickers ignore typed text, so the value is set through the jQuery date picker.
        private async Task SetDatePickerAsync(string selector, DateTime date)
        {
            string expected = date.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
            await _page.EvaluateAsync(@"(a) => {
                const el = document.querySelector(a[0]);
                const p = a[1].split('/').map(Number);
                const d = new Date(p[2], p[0] - 1, p[1]);
                if (window.jQuery && window.jQuery.fn.datepicker) {
                    const $el = window.jQuery(el);
                    // the start/end pickers limit each other, which can clear a same-day value: lift the limits first
                    try { $el.datepicker('setStartDate', null); $el.datepicker('setEndDate', null); } catch (e) { }
                    $el.datepicker('update', d);
                }
                if (el.value !== a[1]) el.value = a[1];
                el.dispatchEvent(new Event('change', { bubbles: true }));
            }", new[] { selector, expected });

            string actual = await _page.Locator(selector).InputValueAsync();
            if (actual != expected)
                throw new InvalidOperationException($"Could not set {selector} to {expected} (field shows '{actual}').");
        }

        private async Task WaitForInfoTextToSettleAsync()
        {
            string previous = string.Empty;
            int stable = 0;
            var deadline = DateTime.UtcNow.AddSeconds(60);
            while (DateTime.UtcNow < deadline)
            {
                string current = string.Empty;
                try { current = await _page.Locator("text=/Showing \\d+ to \\d+ of/").First.InnerTextAsync(new() { Timeout = 2000 }); }
                catch (Exception) { }
                if (current.Length > 0 && current == previous) stable++; else stable = 0;
                previous = current;
                if (stable >= 2) return;
                await Task.Delay(1000);
            }
        }

        // Large bold text placed directly above the grid, so it is inside the grid screenshot
        // and can be copied/sent together with the data.
        private async Task ShowBannerAsync(string text, string cssColor)
        {
            await _page.EvaluateAsync(@"(a) => {
                const old = document.getElementById('autoBanner'); if (old) old.remove();
                const div = document.createElement('div');
                div.id = 'autoBanner';
                div.textContent = a[0];
                div.style.cssText = 'display:block;font-size:34px;font-weight:bold;color:' + a[1] + ';padding:10px 6px 14px 6px;margin:0;background:#fff;white-space:nowrap;';
                const grid = [...document.querySelectorAll('table')].find(t => t.querySelectorAll('tbody tr td').length > 0);
                const anchor = grid ? (grid.closest('.dataTables_scroll') || grid) : null;
                if (anchor) anchor.insertAdjacentElement('beforebegin', div); else document.body.prepend(div);
            }", new[] { text, cssColor });
        }

        // Screenshots only the grid (column headers + rows [+ footer total]), not the whole page.
        // Scroll containers around the grid are opened up so every column and row is in the picture.
        // Very tall grids are split into several images (Chromium limits the size of one capture).
        private async Task<List<string>> CaptureGridAsync(string name)
        {
            await _page.EvaluateAsync(@"() => {
                const tables = [...document.querySelectorAll('table')].filter(t => t.querySelectorAll('tbody tr td').length > 0);
                const grid = tables[0];
                const wrapper = grid.closest('.dataTables_scroll') || grid;
                wrapper.setAttribute('data-grid-capture', '1');
                [wrapper, ...wrapper.querySelectorAll('.dataTables_scrollHead, .dataTables_scrollBody, .dataTables_scrollFoot, .dataTables_scrollHeadInner')]
                    .forEach(e => { e.style.setProperty('overflow', 'visible', 'important'); e.style.setProperty('max-height', 'none', 'important'); e.style.setProperty('height', 'auto', 'important'); });
                // a grid wider than its card must not be clipped by any parent either
                for (let p = wrapper.parentElement; p && p !== document.documentElement; p = p.parentElement)
                    p.style.setProperty('overflow', 'visible', 'important');
                window.scrollTo(0, 0);
            }");

            int docWidth = await _page.EvaluateAsync<int>("Math.max(document.documentElement.scrollWidth, document.body.scrollWidth)");
            int docHeight = await _page.EvaluateAsync<int>("Math.max(document.documentElement.scrollHeight, document.body.scrollHeight)");
            await _page.SetViewportSizeAsync(Math.Max(1500, docWidth + 20), Math.Min(docHeight + 20, 8000));
            await Task.Delay(800);

            var rect = await _page.EvaluateAsync<System.Text.Json.JsonElement>(@"() => {
                const w = document.querySelector('[data-grid-capture]');
                const r = w.getBoundingClientRect();
                let left = r.left, top = r.top, right = r.right, bottom = r.bottom;
                w.querySelectorAll('table, th, td').forEach(t => { const tr = t.getBoundingClientRect(); right = Math.max(right, tr.right); bottom = Math.max(bottom, tr.bottom); });
                const b = document.getElementById('autoBanner');   // the text above the grid is part of the picture
                if (b) {
                    const br = b.getBoundingClientRect();
                    const textRight = br.left + b.scrollWidth;
                    left = Math.min(left, br.left); top = Math.min(top, br.top); right = Math.max(right, textRight);
                }
                return { x: left + window.scrollX, y: top + window.scrollY, w: right - left, h: bottom - top };
            }");
            double x = rect.GetProperty("x").GetDouble(), y = rect.GetProperty("y").GetDouble();
            double w = Math.Ceiling(rect.GetProperty("w").GetDouble()), h = Math.Ceiling(rect.GetProperty("h").GetDouble());
            // make the white text banner as wide as the picture so no page background shows beside it
            await _page.EvaluateAsync("(w) => { const b = document.getElementById('autoBanner'); if (b) { b.style.setProperty('width', w + 'px'); b.style.setProperty('box-sizing', 'border-box'); } }", w);

            var files = new List<string>();
            int chunks = (int)Math.Ceiling(h / MaxScreenshotChunkHeight);
            string prefix = Path.Combine(directoryPath, $"{MMAuditPages.GetIndiaToday():MMddyyyy}- {name}");
            for (int i = 0; i < chunks; i++)
            {
                double offset = i * MaxScreenshotChunkHeight;
                string file = chunks == 1 ? prefix + ".png" : $"{prefix} part{i + 1}.png";
                await _page.ScreenshotAsync(new()
                {
                    Path = file,
                    FullPage = true,
                    Clip = new Clip { X = (float)x, Y = (float)(y + offset), Width = (float)w, Height = (float)Math.Min(MaxScreenshotChunkHeight, h - offset) }
                });
                files.Add(file);
            }
            await _page.SetViewportSizeAsync(1500, 1080);
            Console.WriteLine($"Grid screenshot captured for '{name}' ({files.Count} image(s), {w}x{h}px).");
            return files;
        }

        private static void DeleteScreenshotFiles(IEnumerable<string> files)
        {
            foreach (var file in files)
            {
                // the file can stay locked for a moment after the workbook is saved (indexer/antivirus), so retry
                for (int attempt = 1; attempt <= 5; attempt++)
                {
                    try
                    {
                        if (File.Exists(file)) File.Delete(file);
                        break;
                    }
                    catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                    {
                        if (attempt == 5) Console.WriteLine($"Could not delete screenshot file {file}: {ex.Message}");
                        else Thread.Sleep(1000);
                    }
                }
            }
        }

        // Adds the grid image(s) (which already contain the displayed text) under what the sheet already holds.
        private void AppendSectionToSheet(string heading, List<string> imagePaths, System.Drawing.Color headingColor)
        {
            try
            {
                using var package = new ExcelPackage(new FileInfo(auditWorkbookPath));
                var ws = package.Workbook.Worksheets.FirstOrDefault(w => w.Name == AgingSheetName)
                         ?? package.Workbook.Worksheets.Add(AgingSheetName);
                ws.View.ShowGridLines = false;

                int row = _nextExcelRow > 0
                    ? _nextExcelRow
                    : (ws.Drawings.Count > 0 ? ws.Drawings.Max(d => d.To.Row) + 3 : 1);

                // The text is part of the picture itself, so nothing is typed into the cells.
                int anchor = row - 1; // zero-based row index of the first free row
                foreach (var imagePath in imagePaths)
                {
                    var pic = ws.Drawings.AddPicture($"Section{DateTime.Now.Ticks}_{anchor}", new FileInfo(imagePath));
                    pic.SetPosition(anchor, 0, 0, 0);
                    anchor += (int)Math.Ceiling(pic.Image.Bounds.Height / (double)ExcelRowHeightPixels);
                }
                _nextExcelRow = anchor + 3;
                package.Save();
            }
            catch (Exception ex) when (ex is IOException || ex.ToString().Contains("being used by another process"))
            {
                throw new IOException($"Could not write the screenshot into '{auditWorkbookPath}'. Close the workbook if it is open in Excel and run again. {ex.Message}", ex);
            }
            DeleteScreenshotFiles(imagePaths); // the pictures are now stored inside the workbook
            Console.WriteLine($"Added '{heading}' with its screenshot to sheet '{AgingSheetName}'.");
        }

        private async Task<int> CountBodyRowsAsync()
            => await _page.Locator("table tbody tr:has(td)").CountAsync();

        private async Task WaitForRowsToSettleAsync(int timeoutMs = 60000)
        {
            var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);
            int previous = -1, stableReads = 0;
            while (DateTime.UtcNow < deadline)
            {
                int current = await CountBodyRowsAsync();
                if (current > 0 && current == previous) stableReads++; else stableReads = 0;
                previous = current;
                if (stableReads >= 2) return;
                await Task.Delay(1500);
            }
        }

        private async Task<List<int>> ReadAgingValuesAsync()
        {
            var raw = await _page.EvaluateAsync<string[]>(@"(pat) => {
                const head = [...document.querySelectorAll('th')].find(h => new RegExp(pat, 'i').test(h.textContent));
                if (!head) return [];
                const idx = head.cellIndex;
                const tables = [...document.querySelectorAll('table')].filter(t => t.querySelectorAll('tbody tr td').length > 0);
                const table = tables.find(t => t.contains(head)) || tables[0];
                return [...table.querySelectorAll('tbody tr')].map(tr => tr.cells[idx] ? tr.cells[idx].textContent : '');
            }", AgingColumnPattern);

            var values = new List<int>();
            foreach (var s in raw)
            {
                string digits = Regex.Replace(s ?? string.Empty, @"[^0-9\-]", "");
                if (int.TryParse(digits, out int n)) values.Add(n);
            }
            return values;
        }

        private static bool IsDescending(List<int> values)
        {
            for (int i = 1; i < values.Count; i++)
                if (values[i] > values[i - 1]) return false;
            return true;
        }
    }
}
