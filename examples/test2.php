<?php
$header = ['test_id', 'Acid Uric', 'ALT - GPT', 'AST - GOT', 'BIL', 'BLD', 'Creatinin', 'GLU', 'Glucose', '#GRAN', '#LYM', 'HCT', 'HDL', 'HGB', 'KET', 'LDL', 'LEU', 'MCH', 'MCHC', 'MCV', 'MPV', 'NIT', 'PCT', 'PDW', '%GRAN', '%LYM', '%MID', 'PH', 'PLT', 'PRO', 'RBC', 'RDW', 'SG', 'Triglycerid', 'URE', 'URO', 'WBC', '#mid', 'Blood Test Summary', 'Ure Test Summary', 'serum_biochemistry_summary'];

require_once '../XlsxWriter.php';

try {
  $start = microtime(1);
  $xlsx = new XlsxWriter('test2.xlsx');

  $xlsx->addHeader($header);
  $size = 100000;
  $column = count($header);
  for($i=0;$i<$size;$i++){
    $xlsx->addRow(randRow($column));
    if ($i and $i%100==0)
      echo ($i*100/$size)."\n";
  }

  $xlsx->close();
  echo "Took ".(microtime(1)-$start)." seconds\n";
}catch (Exception $e){
  var_dump($e->getMessage());
}

function randRow($col){
  $row = [];
  while($col){
    $val = rand(0,65000);
    if (rand(0,10)<5)
      $val .= chr(65+$val%26);
    $row[] = $val;
    $col--;
  }
  return $row;
}