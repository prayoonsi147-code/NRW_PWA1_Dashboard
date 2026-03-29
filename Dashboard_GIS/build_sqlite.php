<?php
/**
 * Build SQLite from existing pending Excel cache
 * Run once after upgrading to SQLite:
 *   http://localhost/Claude%20Test%20Cowork/Dashboard_GIS/build_sqlite.php
 *
 * Strategy: Read from Python .cache.json (faster, no PhpSpreadsheet needed)
 * The cache format is: {"mtime": ..., "rows": [[cell, cell, ...], ...]}
 * Data rows start at index 8 (row 9+), header is rows 0-7.
 */

ini_set('memory_limit', '2048M');
set_time_limit(600);

echo "<h2>Build Pending SQLite Database</h2><pre>\n";

// Parse Thai date function
function parse_thai_date_build($val) {
    if ($val instanceof DateTime) {
        $by = $val->format('Y') < 2500 ? $val->format('Y') + 543 : $val->format('Y');
        return [$val, $by];
    }
    if (is_string($val) && strpos($val, '/') !== false) {
        try {
            $parts = explode('/', trim($val));
            if (count($parts) === 3) {
                $dd = (int)$parts[0];
                $mm = (int)$parts[1];
                $yyyy = (int)$parts[2];
                $by = $yyyy > 2500 ? $yyyy : $yyyy + 543;
                $ce_year = $by - 543;
                $dt = new DateTime("$ce_year-$mm-$dd");
                return [$dt, $by];
            }
        } catch (Exception $e) {
            return [null, null];
        }
    }
    return [null, null];
}

$pending_dir = __DIR__ . DIRECTORY_SEPARATOR . 'ข้อมูลดิบ' . DIRECTORY_SEPARATOR . 'ซ่อมท่อค้างระบบ';
if (!is_dir($pending_dir)) {
    echo "ERROR: Pending directory not found: $pending_dir\n";
    exit;
}

// Find Excel files + their cache
$excel_files = [];
foreach (scandir($pending_dir) as $fname) {
    if (preg_match('/\.xlsx?$/i', $fname) && preg_match('/(\d{1,2})-(\d{2})_to_(\d{1,2})-(\d{2})/', $fname)) {
        $excel_files[] = $fname;
    }
}

if (empty($excel_files)) {
    echo "No pending Excel files found in: $pending_dir\n";
    exit;
}

echo "Found " . count($excel_files) . " Excel file(s)\n\n";

foreach ($excel_files as $excel_name) {
    $excel_path = $pending_dir . DIRECTORY_SEPARATOR . $excel_name;
    $cache_path = $excel_path . '.cache.json';  // Python-style cache
    $db_name = preg_replace('/\.xlsx?$/i', '.sqlite', $excel_name);
    $db_path = $pending_dir . DIRECTORY_SEPARATOR . $db_name;

    echo "Processing: $excel_name\n";
    echo "  Excel size: " . round(filesize($excel_path) / 1024 / 1024, 1) . " MB\n";

    // Check if SQLite already exists
    if (file_exists($db_path)) {
        echo "  SQLite already exists: $db_name (skipping, delete it to rebuild)\n\n";
        continue;
    }

    // Check for Python cache
    if (!file_exists($cache_path)) {
        echo "  ERROR: No .cache.json found at: $cache_path\n";
        echo "  Please run the Python server once to generate the cache, or upload files via the web UI.\n\n";
        continue;
    }

    $cache_size = filesize($cache_path);
    echo "  Cache size: " . round($cache_size / 1024 / 1024, 1) . " MB\n";
    echo "  Reading cache JSON...\n";
    flush();

    $start_time = microtime(true);

    try {
        // Read cache JSON
        $json_raw = file_get_contents($cache_path);
        if ($json_raw === false) {
            echo "  ERROR: Could not read cache file\n\n";
            continue;
        }

        echo "  Parsing JSON (" . round(strlen($json_raw) / 1024 / 1024, 1) . " MB)...\n";
        flush();

        $cached = json_decode($json_raw, true);
        unset($json_raw); // Free raw string memory

        if (!$cached || !isset($cached['rows'])) {
            echo "  ERROR: Invalid cache format (no 'rows' key)\n\n";
            continue;
        }

        $rows = $cached['rows'];
        unset($cached); // Free parsed cache memory
        $total_rows = count($rows);

        echo "  Total rows in cache: $total_rows\n";

        // Column indices (0-based)
        $HEADER_ROWS = 8;
        $C_NOTIFY = 2; $C_DATE = 3; $C_FINISH = 5; $C_JOB = 6;
        $C_TYPE = 7; $C_SIDE = 8; $C_TOPIC = 9; $C_DETAIL = 10;
        $C_BRANCH = 19; $C_TEAM = 20; $C_TECH = 21; $C_PIPE = 25; $C_STATUS = 26;

        // Create SQLite database
        $db = new SQLite3($db_path);
        $db->exec('PRAGMA journal_mode=WAL');
        $db->exec('PRAGMA synchronous=NORMAL');

        $db->exec('CREATE TABLE pending_rows (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            row_num INTEGER,
            notify_no TEXT,
            date_val TEXT,
            date_ce TEXT,
            date_by INTEGER,
            month_key TEXT,
            finish_val TEXT,
            finish_ce TEXT,
            job_no TEXT,
            type_val TEXT,
            side TEXT,
            topic TEXT,
            detail TEXT,
            branch TEXT,
            team TEXT,
            tech TEXT,
            pipe TEXT,
            status TEXT
        )');

        $db->exec('CREATE INDEX idx_branch ON pending_rows(branch)');
        $db->exec('CREATE INDEX idx_month_key ON pending_rows(month_key)');
        $db->exec('CREATE INDEX idx_status ON pending_rows(status)');
        $db->exec('CREATE INDEX idx_side ON pending_rows(side)');
        $db->exec('CREATE INDEX idx_job_no ON pending_rows(job_no)');
        $db->exec('CREATE INDEX idx_date_ce ON pending_rows(date_ce)');

        $db->exec('CREATE TABLE meta (key TEXT PRIMARY KEY, value TEXT)');

        $stmt = $db->prepare('INSERT INTO pending_rows
            (row_num, notify_no, date_val, date_ce, date_by, month_key,
             finish_val, finish_ce, job_no, type_val, side, topic, detail,
             branch, team, tech, pipe, status)
            VALUES (:row_num, :notify_no, :date_val, :date_ce, :date_by, :month_key,
                    :finish_val, :finish_ce, :job_no, :type_val, :side, :topic, :detail,
                    :branch, :team, :tech, :pipe, :status)');

        $db->exec('BEGIN TRANSACTION');

        $row_count = 0;
        $skip_count = 0;
        $last_report_dt = null;

        echo "  Inserting data rows (starting from row $HEADER_ROWS)...\n";
        flush();

        for ($r = $HEADER_ROWS; $r < $total_rows; $r++) {
            $row = $rows[$r];

            $date_val = isset($row[$C_DATE]) ? $row[$C_DATE] : null;
            if (!$date_val) { $skip_count++; continue; }

            [$dt, $by] = parse_thai_date_build($date_val);
            if (!$dt || !$by) { $skip_count++; continue; }

            $date_ce = $dt->format('Y-m-d');
            $yy = $by % 100;
            $mm = (int)$dt->format('m');
            $month_key = sprintf('%02d-%02d', $yy, $mm);

            if ($last_report_dt === null || $dt > $last_report_dt) {
                $last_report_dt = $dt;
            }

            // Parse finish date
            $finish_val = (string)(isset($row[$C_FINISH]) ? ($row[$C_FINISH] ?? '') : '');
            $finish_ce = '';
            if ($finish_val) {
                [$fdt, $_] = parse_thai_date_build($finish_val);
                if ($fdt) {
                    $finish_ce = $fdt->format('Y-m-d');
                } elseif (is_string($finish_val) && strlen($finish_val) >= 10) {
                    [$fdt, $_] = parse_thai_date_build(substr($finish_val, 0, 10));
                    if ($fdt) $finish_ce = $fdt->format('Y-m-d');
                }
            }

            $branch = trim((string)(isset($row[$C_BRANCH]) ? $row[$C_BRANCH] : ''));
            if (!$branch) { $skip_count++; continue; }

            $stmt->bindValue(':row_num', $r, SQLITE3_INTEGER);
            $stmt->bindValue(':notify_no', (string)(isset($row[$C_NOTIFY]) ? $row[$C_NOTIFY] : ''), SQLITE3_TEXT);
            $stmt->bindValue(':date_val', is_string($date_val) ? $date_val : $dt->format('d/m/Y'), SQLITE3_TEXT);
            $stmt->bindValue(':date_ce', $date_ce, SQLITE3_TEXT);
            $stmt->bindValue(':date_by', $by, SQLITE3_INTEGER);
            $stmt->bindValue(':month_key', $month_key, SQLITE3_TEXT);
            $stmt->bindValue(':finish_val', $finish_val, SQLITE3_TEXT);
            $stmt->bindValue(':finish_ce', $finish_ce, SQLITE3_TEXT);
            $stmt->bindValue(':job_no', trim((string)(isset($row[$C_JOB]) ? $row[$C_JOB] : '')), SQLITE3_TEXT);
            $stmt->bindValue(':type_val', (string)(isset($row[$C_TYPE]) ? $row[$C_TYPE] : ''), SQLITE3_TEXT);
            $stmt->bindValue(':side', (string)(isset($row[$C_SIDE]) ? $row[$C_SIDE] : ''), SQLITE3_TEXT);
            $stmt->bindValue(':topic', (string)(isset($row[$C_TOPIC]) ? $row[$C_TOPIC] : ''), SQLITE3_TEXT);
            $stmt->bindValue(':detail', (string)(isset($row[$C_DETAIL]) ? $row[$C_DETAIL] : ''), SQLITE3_TEXT);
            $stmt->bindValue(':branch', $branch, SQLITE3_TEXT);
            $stmt->bindValue(':team', (string)(isset($row[$C_TEAM]) ? $row[$C_TEAM] : ''), SQLITE3_TEXT);
            $stmt->bindValue(':tech', (string)(isset($row[$C_TECH]) ? $row[$C_TECH] : ''), SQLITE3_TEXT);
            $stmt->bindValue(':pipe', (string)(isset($row[$C_PIPE]) ? $row[$C_PIPE] : ''), SQLITE3_TEXT);
            $stmt->bindValue(':status', (string)(isset($row[$C_STATUS]) ? $row[$C_STATUS] : ''), SQLITE3_TEXT);
            $stmt->execute();
            $stmt->reset();
            $row_count++;

            // Progress every 10000 rows
            if ($row_count % 10000 === 0) {
                echo "  ... $row_count rows inserted\n";
                flush();
            }
        }

        $db->exec('COMMIT');

        // Free rows memory
        unset($rows);

        // Save meta
        $meta = $db->prepare('INSERT INTO meta (key, value) VALUES (:k, :v)');
        $meta->bindValue(':k', 'row_count', SQLITE3_TEXT);
        $meta->bindValue(':v', (string)$row_count, SQLITE3_TEXT);
        $meta->execute();

        if ($last_report_dt) {
            $by_lrd = $last_report_dt->format('Y') + 543;
            $meta->reset();
            $meta->bindValue(':k', 'update_date', SQLITE3_TEXT);
            $meta->bindValue(':v', sprintf('%02d-%02d-%02d',
                $last_report_dt->format('d'), $last_report_dt->format('m'), $by_lrd % 100), SQLITE3_TEXT);
            $meta->execute();
        }

        $meta->reset();
        $meta->bindValue(':k', 'excel_file', SQLITE3_TEXT);
        $meta->bindValue(':v', $excel_name, SQLITE3_TEXT);
        $meta->execute();

        $db->close();

        $elapsed = round(microtime(true) - $start_time, 1);
        $db_size = round(filesize($db_path) / 1024 / 1024, 1);

        echo "\n  DONE!\n";
        echo "  Rows inserted: $row_count\n";
        echo "  Rows skipped: $skip_count\n";
        echo "  SQLite file: $db_name ($db_size MB)\n";
        echo "  Time: {$elapsed}s\n\n";

    } catch (Exception $e) {
        echo "  ERROR: " . $e->getMessage() . "\n\n";
        if (file_exists($db_path)) unlink($db_path);
    }
}

echo "=== Complete ===\n";
echo "You can now use the Dashboard GIS Tab 3 (Tables 4 & 5).\n";
echo "</pre>";
