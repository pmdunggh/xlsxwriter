<?php

class XlsxWriter
{
  protected $tmpPath;
  protected $xlsxPath;
  protected $savePath;
  protected $textCount = 0;
  protected $row = 1;
  protected $zip;
  protected $dictionary;
  protected $sheet;
  protected $col2name;
  protected $columnCount;
  private $textBuffer = '';
  private $rowBuffer = '';
  private $textBufferSize = 0;
  private $rowBufferSize = 0;

  public function __construct($path)
  {
    //Check library
    if (!class_exists('ZipArchive'))
      throw new XlsxWriterException('Chưa cài đặt PHP exention Zip');

    $this->savePath = $path;
    $this->makeTempDir();
    $this->initCoreFile();
    $this->startSheet();
    $this->startDict();
  }

  private function makeTempDir()
  {
    $tmp = sys_get_temp_dir();
    $dir = strval(microtime(1));
    $ds = DIRECTORY_SEPARATOR;
    while (is_dir($tmp . $ds . $dir) || is_file($tmp . $ds . $dir)) {
      usleep(rand(1, 10));
      $dir = strval(microtime(1));
    }
    $this->tmpPath = $tmp . $ds . $dir;
    if (@mkdir($this->tmpPath) === false)
      throw new XlsxWriterException("Không thể tạo thư mục tạm để xử lý");
    $this->xlsxPath = $this->tmpPath . $ds . 'xlsx';
    @mkdir($this->xlsxPath);
    @mkdir($this->xlsxPath . '/_rels');
    @mkdir($this->xlsxPath . '/docProps');
    @mkdir($this->xlsxPath . '/xl');
    @mkdir($this->xlsxPath . '/xl/_rels');
    @mkdir($this->xlsxPath . '/xl/theme');
    @mkdir($this->xlsxPath . '/xl/worksheets');
  }

  private function initCoreFile()
  {
    $tpl = __DIR__.'/template';
    @copy($tpl . '/_rels/.rels', $this->xlsxPath . '/_rels/.rels');
    @copy($tpl . '/[Content_Types].xml', $this->xlsxPath . '/[Content_Types].xml');
    @copy($tpl . '/docProps/app.xml', $this->xlsxPath . '/docProps/app.xml');
    $core = file_get_contents($tpl . '/docProps/core.xml');
    $core = str_replace('@createtime', (new DateTime())->format('Y-m-d\TH:i:s\Z'), $core);
    @file_put_contents($this->xlsxPath . '/docProps/core.xml', $core);
    @copy($tpl . '/xl/_rels/workbook.xml.rels', $this->xlsxPath . '/xl/_rels/workbook.xml.rels');
    @copy($tpl . '/xl/theme/theme1.xml', $this->xlsxPath . '/xl/theme/theme1.xml');
    @copy($tpl . '/xl/styles.xml', $this->xlsxPath . '/xl/styles.xml');
    @copy($tpl . '/xl/workbook.xml', $this->xlsxPath . '/xl/workbook.xml');
    $this->dictionary = $this->xlsxPath . '/xl/sharedStrings.xml';
    $this->sheet = $this->xlsxPath . '/xl/worksheets/sheet1.xml';
  }

  private function startDict()
  {
    @file_put_contents($this->dictionary, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');
  }

  private function closeDict()
  {
    if ($this->textBufferSize) {
      @file_put_contents($this->dictionary, $this->textBuffer, FILE_APPEND);
      $this->textBuffer='';
      $this->textBufferSize=0;
    }
    @file_put_contents($this->dictionary, '</sst>', FILE_APPEND);
  }

  private function addString($text)
  {
    $text = preg_replace('/[\x00-\x1f]/','',htmlspecialchars($text));
    $text = '<si><t>'.$text.'</t></si>';
    $this->textBuffer .= $text;
    $this->textBufferSize += strlen($text);
    if ($this->textBufferSize > 524288) {
      @file_put_contents($this->dictionary, $this->textBuffer, FILE_APPEND);
      $this->textBuffer='';
      $this->textBufferSize=0;
    }
    //@file_put_contents($this->dictionary, $text, FILE_APPEND);
    return $this->textCount++;
  }

  private function startSheet()
  {
    @file_put_contents($this->sheet, @file_get_contents(__DIR__ . '/template/xl/worksheets/sheet-header.txt'));
  }

  private function closeSheet()
  {
    if ($this->rowBufferSize) {
      @file_put_contents($this->sheet, $this->rowBuffer, FILE_APPEND);
      $this->rowBuffer = '';
      $this->rowBufferSize = 0;
    }
    @file_put_contents($this->sheet, @file_get_contents(__DIR__ . '/template/xl/worksheets/sheet-footer.txt'), FILE_APPEND);
  }

  public function addHeader($row)
  {
    $this->columnCount = count($row);
    $this->initColumnName($this->columnCount);
    $text = "<row r=\"{$this->row}\">";
    foreach ($row as $i => $cell) {
      $strIdx = $this->addString($cell);
      $text .= "<c r=\"{$this->col2name[$i]}{$this->row}\" s=\"1\" t=\"s\"><v>$strIdx</v></c>";
    }
    $text .= '</row>';
    @file_put_contents($this->sheet, $text, FILE_APPEND);
    $this->row++;
  }

  public function addRow($row)
  {
    $text = "<row r=\"{$this->row}\">";
    foreach ($row as $i => $cell) {
      if ($cell === null || $cell === '') continue;
      if (preg_match('/^0\d+$/', $cell)) { //is serial as number
        $strIdx = $this->addString($cell);
        $text .= "<c r=\"{$this->col2name[$i]}{$this->row}\" t=\"s\"><v>$strIdx</v></c>";
      }elseif (preg_match('/^-?\d+(\.\d+)?$/', $cell)) { //is number
        $text .= "<c r=\"{$this->col2name[$i]}{$this->row}\"><v>$cell</v></c>";
      } else { //is text
        $strIdx = $this->addString($cell);
        $text .= "<c r=\"{$this->col2name[$i]}{$this->row}\" t=\"s\"><v>$strIdx</v></c>";
      }
    }
    $text .= '</row>';
    $this->rowBuffer .= $text;
    $this->rowBufferSize += strlen($text);
    if ($this->rowBufferSize > 524288) {
      @file_put_contents($this->sheet, $this->rowBuffer, FILE_APPEND);
      $this->rowBuffer = '';
      $this->rowBufferSize = 0;
    }
    //@file_put_contents($this->sheet, $text, FILE_APPEND);
    $this->row++;
  }

  /**
   * @param int $num  Column index based zero
   * @return string
   */
  public static function convertIndex2Name($num) {
    $numeric = $num % 26;
    $letter = chr(65 + $numeric);
    $num2 = intval($num / 26);
    if ($num2 > 0) {
      return self::convertIndex2Name($num2-1) . $letter;
    } else {
      return $letter;
    }
  }
  private function initColumnName($columnCount)
  {
    $this->col2name = [];
    for ($i = 0; $i < $columnCount; $i++) {
      $this->col2name[] = $this->convertIndex2Name($i);
    }
  }

  public function close()
  {
    $this->closeDict();
    $this->closeSheet();
    $this->compressXlsx();
    $this->removeDir($this->tmpPath);
  }
  private function compressXlsx()
  {
    $zip = new ZipArchive();
    $ret = $zip->open($this->savePath, ZipArchive::CREATE | ZipArchive::OVERWRITE);
    if ($ret !== TRUE)
      throw new XlsxWriterException('Không thể tạo file ' . $this->savePath);

    $rootPath = $this->xlsxPath;
    $files = new RecursiveIteratorIterator(
      new RecursiveDirectoryIterator($rootPath),
      RecursiveIteratorIterator::LEAVES_ONLY
    );

    foreach ($files as $name => $file){
      // Skip directories (they would be added automatically)
      if (!$file->isDir()){
        // Get real and relative path for current file
        $filePath = $file->getRealPath();
        $relativePath = substr($filePath, strlen($rootPath) + 1);

        // Add current file to archive
        $zip->addFile($filePath, $relativePath);
      }
    }
    $zip->close();
  }
  private function removeDir($dir) {
    if (is_dir($dir)) {
      $objects = scandir($dir);
      foreach ($objects as $object) {
        if ($object != "." && $object != "..") {
          $path = $dir.DIRECTORY_SEPARATOR.$object;
          if (is_dir($path))
            $this->removeDir($path);
          else
            @unlink($path);
        }
      }
      @rmdir($dir);
    }
  }
}

class XlsxWriterException extends Exception
{
}