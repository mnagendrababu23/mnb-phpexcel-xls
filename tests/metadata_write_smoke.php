<?php

declare(strict_types=1);

$root = dirname(__DIR__);
$workspace = dirname($root);
spl_autoload_register(static function (string $class) use ($root, $workspace): void {
    $prefix = 'Mnb\\PHPExcel\\';
    if (!str_starts_with($class, $prefix)) return;
    $relative = str_replace('\\', '/', substr($class, strlen($prefix))) . '.php';
    foreach ([$root . '/src/', $workspace . '/mnb-phpexcel-core/src/'] as $base) {
        if (is_file($base . $relative)) { require $base . $relative; return; }
    }
});

use Mnb\PHPExcel\Compound\CompoundFileReader;
use Mnb\PHPExcel\Compound\CompoundFileRewriter;
use Mnb\PHPExcel\Compound\CompoundFileWriter;
use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Core\WorksheetData;
use Mnb\PHPExcel\Format\Xls;
use Mnb\PHPExcel\Support\MnbExcelException;

function check(bool $condition, string $message): void {
    if (!$condition) throw new RuntimeException($message);
}

$runtime = __DIR__ . '/runtime';
@mkdir($runtime, 0775, true);
$source = $runtime . '/metadata-write-source.xls';
$updated = $runtime . '/metadata-write-updated.xls';
$removed = $runtime . '/metadata-write-removed.xls';

Xls::write(new WorkbookData([
    new WorksheetData('Visible', [['A', 'B'], [1, 2]]),
    new WorksheetData('Second', [['X'], [3]]),
], [
    'title' => 'Original title',
    'subject' => 'Original subject',
    'creator' => 'Original Author',
    'last_modified_by' => 'Original Saver',
    'company' => 'Original Company',
    'manager' => 'Original Manager',
    'category' => 'Original Category',
    'custom_properties' => [
        'Project ID' => 'P-1',
        'Approved' => true,
        'Score' => 4.5,
        'Review Date' => ['value' => '2026-08-03T09:30:00Z', 'type' => 'datetime'],
    ],
]), $source);

$initial = Xls::metaInfo($source, ['profile' => 'full']);
check(($initial['document']['title'] ?? null) === 'Original title', 'Native XLS writer did not persist title.');
check(($initial['application']['company'] ?? null) === 'Original Company', 'Native XLS writer did not persist company.');
check(($initial['custom_properties']['count'] ?? null) === 4, 'Native XLS writer did not persist custom properties.');
check(($initial['calculation']['mode'] ?? null) === 'automatic', 'Native XLS writer did not emit calculation metadata.');

Xls::updateMetaInfo($source, $updated, [
    'document' => [
        'title' => 'Updated title',
        'subject' => null,
        'creator' => 'Updated Author',
        'category' => 'Finance',
    ],
    'revision' => [
        'last_saved_by' => 'Updated Saver',
        'revision_number' => '7',
        'total_editing_time_seconds' => 125,
        'document_modified_at' => '2026-08-02T20:00:00Z',
    ],
    'application' => [
        'application_name' => 'MNB PHPExcel XLS',
        'application_version' => 2,
        'manager' => 'New Manager',
        'company' => 'New Company',
    ],
    'custom_properties' => [
        'Project ID' => 'P-2',
        'Approved' => null,
        'Count' => ['value' => 42, 'type' => 'integer'],
        'Released' => ['value' => false, 'type' => 'boolean'],
    ],
    'workbook' => [
        'active_sheet' => 'Second',
        'sheet_visibility' => ['Visible' => 'hidden'],
        'date1904' => true,
    ],
    'calculation' => [
        'mode' => 'manual',
        'iterate' => true,
        'iterate_count' => 25,
        'iterate_delta' => 0.01,
        'calc_on_save' => false,
        'reference_mode' => 'r1c1',
    ],
]);

$meta = Xls::metaInfo($updated, ['profile' => 'full']);
check(($meta['document']['title'] ?? null) === 'Updated title', 'Title update failed.');
check(!array_key_exists('subject', $meta['document']), 'Subject removal failed.');
check(($meta['document']['creator'] ?? null) === 'Updated Author', 'Creator update failed.');
check(($meta['document']['category'] ?? null) === 'Finance', 'Category update failed.');
check(($meta['revision']['last_saved_by'] ?? null) === 'Updated Saver', 'Last saver update failed.');
check(($meta['revision']['revision_number'] ?? null) === '7', 'Revision update failed.');
check(($meta['revision']['total_editing_time'] ?? null) === 125, 'Editing time update failed.');
check(($meta['application']['company'] ?? null) === 'New Company', 'Company update failed.');
check(($meta['application']['manager'] ?? null) === 'New Manager', 'Manager update failed.');
check(($meta['application']['application_version'] ?? null) === '2.0', 'Application version update failed.');
check(($meta['workbook']['active_sheet_name'] ?? null) === 'Second', 'Active sheet update failed.');
check(($meta['workbook']['date_system'] ?? null) === 1904, 'Date system update failed.');
check(($meta['hidden_content']['hidden_sheet_count'] ?? null) === 1, 'Sheet visibility update failed.');
check(($meta['calculation']['mode'] ?? null) === 'manual', 'Calculation mode update failed.');
check(($meta['calculation']['iteration_enabled'] ?? null) === true, 'Iteration update failed.');
check(($meta['calculation']['maximum_iterations'] ?? null) === 25, 'Iteration count update failed.');
check(abs((float) ($meta['calculation']['maximum_change'] ?? -1) - 0.01) < 0.0000001, 'Iteration delta update failed.');
check(($meta['calculation']['save_recalculation'] ?? null) === false, 'Save recalculation update failed.');
check(strtolower((string) ($meta['calculation']['reference_mode'] ?? '')) === 'r1c1', 'Reference mode update failed.');
$custom = [];
foreach ($meta['custom_properties']['items'] as $item) $custom[$item['name']] = $item['value'];
check(($custom['Project ID'] ?? null) === 'P-2', 'Custom property replacement failed.');
check(!array_key_exists('Approved', $custom), 'Custom property deletion failed.');
check(($custom['Count'] ?? null) === 42, 'Integer custom property failed.');
check(($custom['Released'] ?? null) === false, 'Boolean custom property failed.');
check(($custom['Score'] ?? null) === 4.5, 'Untouched custom property was not preserved.');
$updatedRows = Xls::read($updated)->sheet('Second')->toArray();
check(($updatedRows[1][0] ?? null) === 3, 'Workbook cell data changed during metadata update.');

// Same-path update must be atomic and readable.
Xls::updateMetaInfo($updated, $updated, ['document' => ['title' => 'Same path title']]);
check((Xls::metaInfo($updated)['document']['title'] ?? null) === 'Same path title', 'Same-path atomic update failed.');

// Unknown streams and the Workbook stream must remain byte-identical when only properties change.
$richSource = __DIR__ . '/fixtures/rich-metadata.xls';
$richUpdated = $runtime . '/rich-metadata-updated.xls';
$beforeReader = new CompoundFileReader($richSource);
$before = [];
foreach ($beforeReader->directoryEntries() as $entry) {
    if ($entry->type === 2) $before[$entry->name] = hash('sha256', $beforeReader->readStreamById($entry->id));
}
Xls::updateMetaInfo($richSource, $richUpdated, ['document' => ['title' => 'Preserved Package']]);
$afterReader = new CompoundFileReader($richUpdated);
$after = [];
foreach ($afterReader->directoryEntries() as $entry) {
    if ($entry->type === 2) $after[$entry->name] = hash('sha256', $afterReader->readStreamById($entry->id));
}
foreach ($before as $name => $hash) {
    if (in_array($name, ["\x05SummaryInformation", "\x05DocumentSummaryInformation"], true)) continue;
    check(($after[$name] ?? null) === $hash, 'Unrelated OLE stream changed: ' . bin2hex($name));
}
check((Xls::metaInfo($richUpdated)['document']['title'] ?? null) === 'Preserved Package', 'Rich XLS update failed.');

Xls::removePersonalInfo($updated, $removed);
$clean = Xls::metaInfo($removed, ['profile' => 'full']);
check(!array_key_exists('creator', $clean['document']), 'Creator was not removed.');
check(!array_key_exists('last_saved_by', $clean['revision']), 'Last saver was not removed.');
check(!array_key_exists('manager', $clean['application']), 'Manager was not removed.');
check(!array_key_exists('company', $clean['application']), 'Company was not removed.');
check(($clean['custom_properties']['count'] ?? -1) === 0, 'Custom properties were not removed.');

// Property streams must be addable to a valid workbook that does not have them yet.
$sourceCompound = new CompoundFileReader($source);
$minimal = $runtime . '/metadata-no-properties.xls';
file_put_contents($minimal, CompoundFileWriter::build('Workbook', $sourceCompound->readStream('Workbook')));
$minimalUpdated = $runtime . '/metadata-no-properties-updated.xls';
Xls::updateMetaInfo($minimal, $minimalUpdated, [
    'document' => ['title' => 'Added properties'],
    'application' => ['company' => 'Added Company'],
    'custom_properties' => ['Added' => true],
]);
$minimalMeta = Xls::metaInfo($minimalUpdated, ['profile' => 'full']);
check(($minimalMeta['document']['title'] ?? null) === 'Added properties', 'Missing SummaryInformation stream was not added.');
check(($minimalMeta['application']['company'] ?? null) === 'Added Company', 'Missing DocumentSummaryInformation stream was not added.');
check(($minimalMeta['custom_properties']['count'] ?? null) === 1, 'Missing custom-property section was not added.');

// Digital-signature streams must never be invalidated silently.
$signed = $runtime . '/signed-like.xls';
$signedReader = new CompoundFileReader($source);
$signedEntries = $signedReader->directoryEntries();
$signedStreams = [];
foreach ($signedEntries as $entry) {
    if ($entry->type === 2) $signedStreams[$entry->id] = $signedReader->readStreamById($entry->id);
}
[$signedEntries, $signedStreams] = CompoundFileRewriter::upsertRootStreams(
    $signedEntries,
    $signedStreams,
    ['MsoDigitalSignatureEx' => 'test-signature-bytes']
);
file_put_contents($signed, CompoundFileRewriter::build($signedEntries, $signedStreams));
$signedMeta = Xls::metaInfo($signed, ['profile' => 'full']);
check(($signedMeta['security']['digital_signature_present'] ?? null) === true, 'Digital signature stream was not reported.');
$signatureRejected = false;
try {
    Xls::updateMetaInfo($signed, $runtime . '/signed-rejected.xls', ['document' => ['title' => 'No']]);
} catch (MnbExcelException) {
    $signatureRejected = true;
}
check($signatureRejected, 'Signed XLS metadata update was not rejected by default.');
Xls::updateMetaInfo(
    $signed,
    $runtime . '/signed-explicit.xls',
    ['document' => ['title' => 'Explicit']],
    ['allow_invalidate_digital_signatures' => true]
);
check((Xls::metaInfo($runtime . '/signed-explicit.xls')['document']['title'] ?? null) === 'Explicit', 'Explicit signed-file update failed.');

$guard = $runtime . '/atomic-guard.xls';
file_put_contents($guard, 'unchanged');
try {
    Xls::updateMetaInfo($source, $guard, ['revision' => ['document_modified_at' => 'not-a-date']]);
    throw new RuntimeException('Invalid date did not fail.');
} catch (MnbExcelException) {
    check(file_get_contents($guard) === 'unchanged', 'Failed update replaced the destination.');
}

$libreOfficeChecked = false;
$lo = trim((string) shell_exec('command -v libreoffice 2>/dev/null'));
if ($lo !== '') {
    $loDir = $runtime . '/lo-check';
    @mkdir($loDir, 0775, true);
    exec(escapeshellarg($lo) . ' --headless --convert-to xlsx --outdir ' . escapeshellarg($loDir) . ' ' . escapeshellarg($updated) . ' 2>&1', $output, $status);
    check($status === 0 && is_file($loDir . '/' . pathinfo($updated, PATHINFO_FILENAME) . '.xlsx'), 'LibreOffice rejected updated XLS: ' . implode("\n", $output));
    $libreOfficeChecked = true;
}

echo json_encode([
    'status' => 'passed',
    'updated_title' => Xls::metaInfo($updated)['document']['title'] ?? null,
    'custom_properties' => $meta['custom_properties']['count'],
    'libreoffice_checked' => $libreOfficeChecked,
], JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES) . PHP_EOL;
