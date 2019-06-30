<?php
$header = ['test_id', 'Acid Uric', 'ALT - GPT', 'AST - GOT', 'BIL', 'BLD', 'Creatinin'];
$data =[
  ['00001', '211', '16', '17', 'Âm tính', '2+', '61'],
  ['00002', '182', '14', '16', 'Âm tính', 'Âm tính', '60'],
  ['00003', '394', '18', '20', 'Âm tính', 'Âm tính', '78'],
  ['00004', '256', '132', '73', 'Âm tính', 'Âm tính', '72'],
  ['00005', '340', '28', '22', 'Âm tính', 'Âm tính', '72']
];

require_once '../XlsxWriter.php';

try {
  $xlsx = new XlsxWriter('test1.xlsx');

  $xlsx->addHeader($header);
  foreach ($data as $row)
    $xlsx->addRow($row);

  $xlsx->close();
}catch (Exception $e){
  var_dump($e->getMessage());
}

