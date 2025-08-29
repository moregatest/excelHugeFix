# Excel 分析器 PHP 整合指引

## 概述

本文件說明如何在 PHP 專案中整合 Excel 分析器工具，用於檢測和修復 Excel 檔案的尺寸問題。

## 系統需求

- **Python**: 3.7.16 或更高版本
- **uv**: Python 套件管理工具
- **PHP**: 7.0 或更高版本
- **系統**: macOS, Linux, Windows

## 安裝與設置

### 1. 安裝 uv（如果尚未安裝）

```bash
# macOS/Linux
curl -LsSf https://astral.sh/uv/install.sh | sh

# 或使用 brew
brew install uv
```

### 2. 驗證工具可用性

```bash
cd /path/to/excel_analyzer
uv run excel_analyzer_cli.py --version
```

## 退出碼說明

Excel 分析器使用標準退出碼與 PHP 程式通信：

| 退出碼 | 狀態 | 說明 |
|--------|------|------|
| `0` | 正常 | Excel 檔案無問題 |
| `1` | 有問題 | 檢測到問題（檢測模式）或已修復（修復模式） |
| `2` | 錯誤 | 分析失敗（檔案不存在、格式錯誤等） |

## PHP 整合方式

### 方法一：使用 `exec()` 函數（推薦）

```php
<?php

class ExcelAnalyzer
{
    private $analyzerPath;
    
    public function __construct($analyzerPath = '/path/to/excel_analyzer')
    {
        $this->analyzerPath = rtrim($analyzerPath, '/');
    }
    
    /**
     * 檢查 Excel 檔案是否正常
     * 
     * @param string $filePath Excel 檔案路徑
     * @return array 分析結果
     */
    public function checkExcelFile($filePath)
    {
        // 驗證檔案存在
        if (!file_exists($filePath)) {
            return [
                'status' => 'error',
                'code' => 2,
                'message' => '檔案不存在',
                'file_path' => $filePath
            ];
        }
        
        // 建構命令
        $command = sprintf(
            'cd %s && uv run excel_analyzer_cli.py --check %s 2>&1',
            escapeshellarg($this->analyzerPath),
            escapeshellarg($filePath)
        );
        
        // 執行命令
        exec($command, $output, $exitCode);
        
        // 解析結果
        return $this->parseResult($exitCode, $output, $filePath);
    }
    
    /**
     * 修復 Excel 檔案問題
     * 
     * @param string $filePath Excel 檔案路徑
     * @return array 修復結果
     */
    public function fixExcelFile($filePath)
    {
        if (!file_exists($filePath)) {
            return [
                'status' => 'error',
                'code' => 2,
                'message' => '檔案不存在',
                'file_path' => $filePath
            ];
        }
        
        $command = sprintf(
            'cd %s && uv run excel_analyzer_cli.py --fix %s 2>&1',
            escapeshellarg($this->analyzerPath),
            escapeshellarg($filePath)
        );
        
        exec($command, $output, $exitCode);
        
        return $this->parseResult($exitCode, $output, $filePath);
    }
    
    /**
     * 解析分析結果
     */
    private function parseResult($exitCode, $output, $originalPath)
    {
        switch ($exitCode) {
            case 0:
                return [
                    'status' => 'ok',
                    'code' => 0,
                    'message' => '檔案正常，無問題',
                    'file_path' => $originalPath,
                    'has_issues' => false
                ];
                
            case 1:
                // 從輸出中取得處理後的檔案路徑
                $processedPath = !empty($output) ? trim($output[0]) : $originalPath;
                
                return [
                    'status' => 'issues_found',
                    'code' => 1,
                    'message' => '發現問題並已處理',
                    'file_path' => $processedPath,
                    'has_issues' => true,
                    'output' => $output
                ];
                
            case 2:
            default:
                return [
                    'status' => 'error',
                    'code' => $exitCode,
                    'message' => '分析失敗',
                    'file_path' => $originalPath,
                    'error_output' => $output
                ];
        }
    }
}

// 使用範例
$analyzer = new ExcelAnalyzer('/path/to/excel_analyzer');

// 檢查檔案
$result = $analyzer->checkExcelFile('/path/to/your/file.xlsx');

if ($result['status'] === 'ok') {
    echo "檔案正常\n";
} elseif ($result['status'] === 'issues_found') {
    echo "發現問題，需要修復\n";
    
    // 自動修復
    $fixResult = $analyzer->fixExcelFile('/path/to/your/file.xlsx');
    echo "修復結果: " . $fixResult['message'] . "\n";
} else {
    echo "錯誤: " . $result['message'] . "\n";
}
?>
```

### 方法二：使用 `shell_exec()` 函數（簡化版）

```php
<?php

function checkExcelFile($filePath, $analyzerPath = '/path/to/excel_analyzer')
{
    $command = sprintf(
        'cd %s && uv run excel_analyzer_cli.py --check %s; echo "EXIT_CODE:$?"',
        escapeshellarg($analyzerPath),
        escapeshellarg($filePath)
    );
    
    $output = shell_exec($command);
    
    // 解析退出碼
    if (preg_match('/EXIT_CODE:(\d+)/', $output, $matches)) {
        $exitCode = (int)$matches[1];
        
        switch ($exitCode) {
            case 0:
                return ['status' => 'ok', 'message' => '檔案正常'];
            case 1:
                return ['status' => 'has_issues', 'message' => '發現問題'];
            case 2:
                return ['status' => 'error', 'message' => '分析失敗'];
        }
    }
    
    return ['status' => 'unknown', 'message' => '無法解析結果'];
}

// 使用範例
$result = checkExcelFile('/path/to/file.xlsx');
echo $result['message'] . "\n";
?>
```

## 進階用法

### 批次處理多個檔案

```php
<?php

class BatchExcelAnalyzer extends ExcelAnalyzer
{
    /**
     * 批次檢查多個 Excel 檔案
     * 
     * @param array $filePaths 檔案路徑陣列
     * @return array 批次處理結果
     */
    public function batchCheck(array $filePaths)
    {
        $results = [];
        $summary = [
            'total' => count($filePaths),
            'ok' => 0,
            'issues' => 0,
            'errors' => 0
        ];
        
        foreach ($filePaths as $filePath) {
            $result = $this->checkExcelFile($filePath);
            $results[$filePath] = $result;
            
            // 統計
            switch ($result['status']) {
                case 'ok':
                    $summary['ok']++;
                    break;
                case 'issues_found':
                    $summary['issues']++;
                    break;
                case 'error':
                    $summary['errors']++;
                    break;
            }
        }
        
        return [
            'summary' => $summary,
            'details' => $results
        ];
    }
    
    /**
     * 自動修復有問題的檔案
     */
    public function autoFixIssues(array $filePaths)
    {
        $checkResults = $this->batchCheck($filePaths);
        $fixResults = [];
        
        foreach ($checkResults['details'] as $filePath => $result) {
            if ($result['status'] === 'issues_found') {
                $fixResults[$filePath] = $this->fixExcelFile($filePath);
            }
        }
        
        return $fixResults;
    }
}

// 使用範例
$analyzer = new BatchExcelAnalyzer('/path/to/excel_analyzer');

$files = [
    '/path/to/file1.xlsx',
    '/path/to/file2.xls',
    '/path/to/file3.xlsx'
];

$results = $analyzer->batchCheck($files);

printf("處理完成: %d 個檔案，%d 正常，%d 有問題，%d 錯誤\n", 
    $results['summary']['total'],
    $results['summary']['ok'],
    $results['summary']['issues'],
    $results['summary']['errors']
);

// 自動修復有問題的檔案
$fixResults = $analyzer->autoFixIssues($files);
foreach ($fixResults as $file => $result) {
    echo "修復 $file: " . $result['message'] . "\n";
}
?>
```

### Web 介面整合

```php
<?php
// api/excel-check.php

header('Content-Type: application/json');

try {
    if (!isset($_POST['file']) || !is_uploaded_file($_FILES['file']['tmp_name'])) {
        throw new Exception('未選擇檔案');
    }
    
    $uploadFile = $_FILES['file'];
    $tempPath = $uploadFile['tmp_name'];
    
    // 驗證檔案類型
    $allowedTypes = ['application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
    if (!in_array($uploadFile['type'], $allowedTypes)) {
        throw new Exception('不支援的檔案格式');
    }
    
    // 分析檔案
    $analyzer = new ExcelAnalyzer('/path/to/excel_analyzer');
    $result = $analyzer->checkExcelFile($tempPath);
    
    // 回傳 JSON 結果
    echo json_encode([
        'success' => true,
        'data' => $result
    ]);
    
} catch (Exception $e) {
    http_response_code(400);
    echo json_encode([
        'success' => false,
        'error' => $e->getMessage()
    ]);
}
?>
```

## 錯誤處理與除錯

### 常見問題

1. **命令找不到**
   ```php
   // 確保路徑正確
   $command = 'cd /correct/path && uv run excel_analyzer_cli.py --check file.xlsx';
   ```

2. **權限問題**
   ```bash
   chmod +x excel_analyzer_cli.py
   ```

3. **Python 環境問題**
   ```bash
   # 檢查 uv 是否可用
   which uv
   uv --version
   ```

### 除錯模式

```php
<?php

class DebugExcelAnalyzer extends ExcelAnalyzer
{
    private $debug = false;
    
    public function enableDebug($enable = true)
    {
        $this->debug = $enable;
    }
    
    public function checkExcelFile($filePath)
    {
        $startTime = microtime(true);
        
        $command = sprintf(
            'cd %s && uv run excel_analyzer_cli.py %s %s 2>&1',
            escapeshellarg($this->analyzerPath),
            $this->debug ? '--debug' : '--check',
            escapeshellarg($filePath)
        );
        
        if ($this->debug) {
            echo "執行命令: $command\n";
        }
        
        exec($command, $output, $exitCode);
        
        $endTime = microtime(true);
        $duration = round(($endTime - $startTime) * 1000, 2);
        
        if ($this->debug) {
            echo "執行時間: {$duration}ms\n";
            echo "退出碼: $exitCode\n";
            echo "輸出:\n" . implode("\n", $output) . "\n";
        }
        
        return $this->parseResult($exitCode, $output, $filePath);
    }
}
?>
```

## 效能考量

### 1. 非同步處理

```php
<?php

class AsyncExcelAnalyzer
{
    /**
     * 非同步檢查檔案（背景執行）
     */
    public function checkAsync($filePath, $callbackUrl = null)
    {
        $command = sprintf(
            'cd %s && nohup uv run excel_analyzer_cli.py --check %s > /tmp/excel_check_%s.log 2>&1 & echo $!',
            escapeshellarg($this->analyzerPath),
            escapeshellarg($filePath),
            md5($filePath)
        );
        
        $pid = shell_exec($command);
        
        return [
            'pid' => trim($pid),
            'log_file' => '/tmp/excel_check_' . md5($filePath) . '.log'
        ];
    }
    
    /**
     * 檢查非同步任務狀態
     */
    public function checkAsyncStatus($pid)
    {
        $command = "ps -p $pid -o pid=";
        $result = shell_exec($command);
        
        return !empty(trim($result));
    }
}
?>
```

### 2. 快取機制

```php
<?php

class CachedExcelAnalyzer extends ExcelAnalyzer
{
    private $cacheDir;
    
    public function __construct($analyzerPath, $cacheDir = '/tmp/excel_cache')
    {
        parent::__construct($analyzerPath);
        $this->cacheDir = $cacheDir;
        
        if (!is_dir($cacheDir)) {
            mkdir($cacheDir, 0755, true);
        }
    }
    
    public function checkExcelFile($filePath)
    {
        // 使用檔案修改時間和大小作為快取鍵
        $fileHash = md5($filePath . filemtime($filePath) . filesize($filePath));
        $cacheFile = $this->cacheDir . '/check_' . $fileHash . '.json';
        
        // 檢查快取
        if (file_exists($cacheFile) && (time() - filemtime($cacheFile)) < 3600) {
            return json_decode(file_get_contents($cacheFile), true);
        }
        
        // 執行檢查
        $result = parent::checkExcelFile($filePath);
        
        // 儲存快取
        file_put_contents($cacheFile, json_encode($result));
        
        return $result;
    }
}
?>
```

## 部署建議

1. **環境變數設定**
   ```bash
   # .env 檔案
   EXCEL_ANALYZER_PATH=/path/to/excel_analyzer
   EXCEL_ANALYZER_CACHE_DIR=/var/cache/excel_analyzer
   ```

2. **日誌記錄**
   ```php
   // 記錄所有 Excel 分析操作
   error_log("Excel分析: $filePath, 結果: " . $result['status']);
   ```

3. **監控與警報**
   ```php
   // 當發現大量問題檔案時發送警報
   if ($summary['issues'] > $summary['total'] * 0.5) {
       // 發送警報通知
       mail('admin@example.com', 'Excel檔案問題警報', '發現大量問題檔案');
   }
   ```

## 結論

通過本指引，您可以輕鬆地將 Excel 分析器整合到 PHP 專案中，實現自動化的 Excel 檔案品質檢測和修復功能。建議在生產環境中使用前進行充分測試，並根據實際需求調整配置。