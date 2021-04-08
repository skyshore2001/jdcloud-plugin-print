<?php

/*! jdcloud-plugin-url BEGIN */
function api_getShareUrl()
{
	$param = [
		"get" => $_GET,
		"post" => $_POST,
		"ses" => session_id()
	];
	$p = json_encode($param, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
	//addLog($p);
	return makeUrl(getBaseUrl() . "url.php", ["p"=>jdEncrypt($p)]);
}
/*! jdcloud-plugin-url END */

function printFile($fmt, $ret, $fname, $tpl)
{
	if ($fmt === "excel") {
		header("Content-disposition: attachment; filename=" . $fname . ".xlsx");
		header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		header("Content-Transfer-Encoding: binary");
	
		$retArr = table2objarr($ret);
	
		require_once 'class/vendor/phpoffice/phpexcel/Classes/PHPExcel.php';
		require_once 'class/vendor/phpoffice/phpexcel/Classes/PHPExcel/IOFactory.php';

		// 模板
		$file = "./template/$tpl.xlsx";
		if (!file_exists($file)) {
			throw new MyException(E_PARAM, "template file: $file not exist");
		}
		//logit("start load excel file");

		$reader = PHPExcel_IOFactory::createReader("Excel2007");
		$spreadsheet = $reader->load($file);

		$sheet = $spreadsheet->getActiveSheet();
		$writer = PHPExcel_IOFactory::createWriter($spreadsheet, 'Excel2007');

		// Get the highest row and column numbers referenced in the worksheet
		$highestRow = $sheet->getHighestRow(); // e.g. 10
		$highestColumn = $sheet->getHighestColumn(); // e.g 'F'
		$highestColumnIndex = PHPExcel_Cell::columnIndexFromString($highestColumn);

		logit("maxrow: $highestRow; maxcol: $highestColumn");
		
		if (isset($_GET["align"])) {
			$align = PHPExcel_Style_Alignment::HORIZONTAL_LEFT;
			switch($_GET["align"]) {
				case middle:
					$align = PHPExcel_Style_Alignment::HORIZONTAL_MIDDLE;
					break;
				case right:
					$align = PHPExcel_Style_Alignment::HORIZONTAL_RIGHT;
					break;
				default:
					break;
			};

			$styleArray = [
				'alignment' => [
				'horizontal' => $align,
				],
			];

			$sheet->getStyle('A1:'.$highestColumn.$highestRow)->applyFromArray($styleArray);
		}

		if (count($retArr) == 1) {
			$r = $retArr[0];
			logit($r);

			$subRow = 0;
			$subCol = 0;
			$seqCol = 0;
			$subobj = null;

			for ($row = 1; $row <= $highestRow; ++$row) {
				if ($row == $subRow)
					continue;

				//这个循环用于处理主表的数据
				for ($col = 0; $col <= $highestColumnIndex; ++$col) {
					$cellValue = $sheet->getCellByColumnAndRow($col, $row)->getValue();

					$match = [];
					if (isset($cellValue)) {
						if (stripos($cellValue, "START_SEQ") !== false) {
							$seqCol = $col;
							logit("start_seq: " . $seqCol);
						}

						$match = preg_replace('/.*{(.+)}.*/iu', '$1', $cellValue);
						logit("matched value: " . $match);

						if (stripos($cellValue, "START_ROW") !== false) {
							$subRow = $row;
							$subCol = $col;
							$subobj = $match;

							logit("subRow: " . $subRow . " subCol:" . $subCol . " subobj: " . $subobj);
							break;
						}

						if ($match != $cellValue) {
							$isImg = (stripos($match, "IMG::") !== false);
							$ret = getField($match, $r, $isImg);

							//生成和设置图片
							if ($isImg) {
								//生成二维码
								include 'class/phpqrcode/phpqrcode.php'; 

								logit("create img: " . $ret);
								$errorCorrectionLevel = 'L';//容错级别
								$matrixPointSize = 6;//生成图片大小
								$imgPath = './qrcode.png';
								QRcode::png($ret, $imgPath, $errorCorrectionLevel, $matrixPointSize, 2);

								$drawing = new PHPExcel_Worksheet_Drawing();
								//$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
								$drawing->setName('qrCode');
								$drawing->setPath($imgPath);

								//col起始位置是A, 就是65
								$colName = chr(65 + $col) . $row;

								$drawing->setHeight(126);
								$drawing->setWidth(126);
								$drawing->setCoordinates($colName);
								$drawing->setOffsetX(15);
								$drawing->setOffsetY(2);
								$drawing->setWorksheet($sheet);
								$sheet->setCellValueByColumnAndRow($col, $row, " ");
							}
							else { //设置文字内容
								logit("set field $col, $row, $ret");
								$sheet->setCellValueByColumnAndRow($col, $row, $ret);
							}
						}
					}
				}
			}

			if (isset($subobj)) {
				$subcnt = count($r[$subobj]);
				//$sheet->insertNewRowBefore(7, 2);
				logit("enter subobj cnt: " . $subcnt);

				//根据子表行数来插入并复制子表第一行的信息
				if ($subcnt > 1) {
					$row = $subRow + 1;
					$sheet->insertNewRowBefore($subRow+1, $subcnt-1);
					for($i=0; $i<$subcnt-1; $i++){
						for ($col = $subCol; $col <= $highestColumnIndex; ++$col) {
							$cellValue = $sheet->getCellByColumnAndRow($col, $subRow)->getValue();

							if (isset($cellValue)) {
								$sheet->setCellValueByColumnAndRow($col, $row, $cellValue);
							}
						}
						$row++;
					}
				}

				$row = $subRow;
				for($i=0; $i<$subcnt; $i++){
					$sub = $r[$subobj][$i];

					for ($col = $subCol; $col <= $highestColumnIndex; ++$col) {
						$cellValue = $sheet->getCellByColumnAndRow($col, $row)->getValue();

						$match = [];
						if ($col == $seqCol) {
							$sheet->setCellValueByColumnAndRow($col, $row, $i+1);
						}
						else if (isset($cellValue)) {
							$match = preg_replace('/.*{(.+)}.*/iu', '$1', $cellValue);
							logit("matched value: " . $match);

							if ($match != $cellValue) {
								$ret = getField($match, $sub);

								logit("set field $col, $row, $ret");
								$sheet->setCellValueByColumnAndRow($col, $row, $ret);
							}
						}
					}

					$row++;
				}
			}

			$writer->save('php://output');
		}

		return true;
	}
}

function getField($raw, $r, $isImg=false) {
	if ($isImg) {
		$raw = str_replace("IMG::","",$raw);
	}

	logit("rawval: " . $raw . " r[rawval]: " . $r[$raw] . " isImg: " . $isImg);

	$expr = preg_replace_callback('/(?<![0-9])\$?([a-z_]\w*\b)(?!\()/iU', function ($ms) use($r, $isImg) {
		$m = $ms[0];
		if ($isImg) {
			$expr = "$m=" . eval("return" . '$r["' . $m . '"];');
		}
		else if ($ms[0] == "FMT_D") {
			$expr = $ms[0];
		}
		else if ($ms[0] == "ROW_START") {
			$expr = $ms[0];
		}
		else {
			$expr = '$r["' . $m . '"]';
		}

		//logit("ms[0]: " . $m . " expr: " . $expr);
		return $expr;
	}, $raw);

	$ret = null;
	if (isset($expr) && strlen($expr) > 0) {
		if ($isImg) {
			$ret = $expr;
		}
		else {
			$expr = "return (" . $expr . ");";
			$ret = eval($expr);
			if (!isset($ret)) {
				$ret = "";
			}
		}
		logit("after eval: " . $ret);
	}

	return $ret;
}

// vi: foldmethod=marker
