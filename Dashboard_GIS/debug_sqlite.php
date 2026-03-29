<?php
ini_set('memory_limit', '256M');
header('Content-Type: text/html; charset=utf-8');
echo "<pre>\n";

$pending_dir = __DIR__ . '/ข้อมูลดิบ/ซ่อมท่อค้างระบบ';
$db_files = glob($pending_dir . '/*.sqlite');

if (empty($db_files)) { echo "No .sqlite files found\n"; exit; }

$db_path = $db_files[0];
echo "DB: " . basename($db_path) . "\n\n";

$db = new SQLite3($db_path, SQLITE3_OPEN_READONLY);

// Total rows
echo "Total rows: " . $db->querySingle("SELECT COUNT(*) FROM pending_rows") . "\n\n";

// Meta
echo "=== META ===\n";
$r = $db->query("SELECT * FROM meta");
while ($row = $r->fetchArray(SQLITE3_ASSOC)) echo "  {$row['key']} = {$row['value']}\n";

// Sample status values
echo "\n=== TOP STATUS VALUES ===\n";
$r = $db->query("SELECT status, COUNT(*) as cnt FROM pending_rows GROUP BY status ORDER BY cnt DESC LIMIT 10");
while ($row = $r->fetchArray(SQLITE3_ASSOC)) echo "  [{$row['status']}] = {$row['cnt']}\n";

// Sample side values
echo "\n=== TOP SIDE VALUES ===\n";
$r = $db->query("SELECT side, COUNT(*) as cnt FROM pending_rows GROUP BY side ORDER BY cnt DESC LIMIT 10");
while ($row = $r->fetchArray(SQLITE3_ASSOC)) echo "  [{$row['side']}] = {$row['cnt']}\n";

// Month keys
echo "\n=== MONTH KEYS ===\n";
$r = $db->query("SELECT month_key, COUNT(*) as cnt FROM pending_rows GROUP BY month_key ORDER BY month_key");
while ($row = $r->fetchArray(SQLITE3_ASSOC)) echo "  {$row['month_key']} = {$row['cnt']}\n";

// Sample job_no
echo "\n=== JOB_NO EMPTY COUNT ===\n";
echo "  NULL or empty: " . $db->querySingle("SELECT COUNT(*) FROM pending_rows WHERE job_no IS NULL OR job_no = ''") . "\n";
echo "  Has value: " . $db->querySingle("SELECT COUNT(*) FROM pending_rows WHERE job_no IS NOT NULL AND job_no != ''") . "\n";

// Latest month
$latest = $db->querySingle("SELECT month_key FROM pending_rows ORDER BY date_ce DESC LIMIT 1");
echo "\n=== LATEST MONTH: $latest ===\n";

// Nojob filter test
echo "\n=== NOJOB FILTER TEST (month=$latest) ===\n";
$sql = "SELECT COUNT(*) FROM pending_rows WHERE month_key = '$latest'";
echo "  All rows in latest month: " . $db->querySingle($sql) . "\n";

$sql .= " AND (side = 'ด้านท่อแตกรั่ว' OR topic LIKE '%ท่อแตก%' OR topic LIKE '%ท่อรั่ว%' OR topic LIKE '%แตกรั่ว%' OR detail LIKE '%ท่อแตก%' OR detail LIKE '%ท่อรั่ว%' OR detail LIKE '%แตกรั่ว%')";
echo "  + pipe filter: " . $db->querySingle($sql) . "\n";

$sql2 = "SELECT COUNT(*) FROM pending_rows WHERE month_key = '$latest' AND (side = 'ด้านท่อแตกรั่ว' OR topic LIKE '%ท่อแตก%' OR topic LIKE '%ท่อรั่ว%' OR topic LIKE '%แตกรั่ว%' OR detail LIKE '%ท่อแตก%' OR detail LIKE '%ท่อรั่ว%' OR detail LIKE '%แตกรั่ว%') AND (job_no IS NULL OR job_no = '')";
echo "  + no job: " . $db->querySingle($sql2) . "\n";

$sql3 = $sql2 . " AND status NOT LIKE '%ดำเนินการแล้วเสร็จ%'";
echo "  + not done: " . $db->querySingle($sql3) . "\n";

// Detail filter test
echo "\n=== DETAIL FILTER TEST (ซ่อมไม่เสร็จ) ===\n";
echo "  status LIKE ซ่อมไม่เสร็จ: " . $db->querySingle("SELECT COUNT(*) FROM pending_rows WHERE status LIKE '%ซ่อมไม่เสร็จ%'") . "\n";

// Sample rows
echo "\n=== SAMPLE 3 ROWS ===\n";
$r = $db->query("SELECT * FROM pending_rows LIMIT 3");
while ($row = $r->fetchArray(SQLITE3_ASSOC)) {
    echo "  ---\n";
    foreach ($row as $k => $v) echo "  $k: $v\n";
}

$db->close();
echo "</pre>";
