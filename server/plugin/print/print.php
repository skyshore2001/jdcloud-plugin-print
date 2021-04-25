<?php

class PrintUtil {
	protected $ret;
	protected $lastCol;
	protected $lastRow;
	protected $sheet;
	protected $obj;
	protected $startRow;
	protected $endRow;
	protected $tempVal;
	public static $curVal;

	public function __construct($sheet, $ret, $lastRow, $lastCol) {
        $this->ret = $ret;
        $this->sheet = $sheet;
        $this->lastCol = $lastCol;
        $this->lastRow = $lastRow;
    }

	public function handleCell ($col, $row, &$lastRow) {
		//获取原始值
		$sheet = $this->sheet;
		$cellValue = $sheet->getCellByColumnAndRow($col, $row)->getValue();
		if (!isset($cellValue)) {
			return;
		}

		addLog("row: $row, col: $col, cell ori value: " . $cellValue);
		$this->cellValue = $cellValue;
		$matched = false;
		$isImg = false;

		//解析
		$expr = preg_replace_callback('/\{([^{}]+)\}/iU', function ($ms) use($col, $row, &$isImg, &$lastRow, &$matched) {
			$m = $ms[0];
			$matched = true;

			addLog("parse first time- ms[0]: " . $ms[0] . " ms[1]: " . $ms[1]);
			//logit("ms[1]: " . $ms[1] . " expr: " . $expr);

			$isTable = false;
			$subObj = null;

			$expr = preg_replace_callback('/(?<![0-9])\$?([a-z_]\w*\b)(?!\()/iU', function ($ms1) use($col, $row, &$lastRow, &$isImg, &$isTable, &$subObj) {
				$expr1 = "";
				switch($ms1[0]) {
					case "j_for":
						$isTable = true;
						break;
					case "j_img":
						$isImg = true;
						break;
					case "FMT_D":
					case "FMT_DT":
						$expr1 = $ms1[0];
						break;
					default:
						if ($isTable)
							$subObj = $ms1[0];
						else if ($isImg){
							$expr1 = $ms1[0];
						}
						else {
							$expr1 = '$r["' . $ms1[0] . '"]';
						}
						break;
				}

				logit("ms1[0]: " . $ms1[0] . " expr: " . $expr1);
				return $expr1;
			}, $ms[1]);

			logit("ms[1]: " . $ms[1] . " expr: " . $expr);

			$ret = null;

			//根据一个{}里的信息可以判断对应的cell, 可以做cell类别和多行的处理了
			if ($isTable) {
				$expr = "";
				$cnt = isset($subObj) ? count($this->ret[$subObj]) : count($this->ret);

				$this->obj = isset($subObj) ? $this->ret[$subObj] : $this->ret;
				$this->startRow = $row;	
				$this->endRow = $row + $cnt;

				logit("start add row");

				//j-for这一个括号里的内容去掉
				$tempVal = preg_replace_callback('/\{([^{}]+)\}/iU', function ($ms2) {
					return "";
				}, $this->cellValue, 1);

				//logit("temp val: " . $tempVal);
				$this->sheet->setCellValueByColumnAndRow($col, $row, $tempVal);
				$this->addRow($cnt, $row, $col);	
				$lastRow += $cnt-1;

				logit("end add row");
			}
			else if (isset($expr) && strlen($expr) > 0 && $matched) {
				if (isset($this->obj))
				{
					if ($row <= $this->endRow) {
						$r = $this->obj[$row - $this->startRow];
					}
					else {
						$this->obj = null;
						$r = $this->ret;
					}
				}
				else {
					$r = $this->ret;
				}

				PrintUtil::$curVal = $r;

				logit("before eval: " . $expr);
				if ($isImg) {
					$imgArr= explode(";", $expr);

					foreach($imgArr as $v) {
						$expr = "return (" . $v . ");";

						logit("full expr: " . $expr);
						if (strlen($ret) > 0) {
							$ret .= ";";
						}
						$ret .= eval($expr);
					}
				}
				else {
					$expr = "return (" . $expr . ");";
					logit("full expr: " . $expr);

					$ret = eval($expr);
				}

				if (!isset($ret)) {
					$ret = "";
				}
				logit("after eval: " . $ret);
			}

			return $ret;
		}, $cellValue);

		if ($matched) {
			if ($isImg) {
				logit("image expr: " . $expr);
				$exprArr = explode(';', $expr);
				$width = @$exprArr[1];
				$height = @$exprArr[2];

				if (!isset($width)) {
					$width = 75;
				}
				if (!isset($height)) {
					$height = 75;
				}
				$this->setImg($exprArr[0], $col, $row, $width, $height); 
			}
			else {
				$sheet->setCellValueByColumnAndRow($col, $row, $expr);
			}
		}
	}

	protected function addRow($cnt, $row, $col) {
		logit($cnt . " - " . $row . " - " . $col);
		if ($cnt <= 1)
			return;

		$newRow = $row + 1;

		logit(000);
		$sheet = $this->sheet;
		$sheet->insertNewRowBefore($newRow, $cnt-1);

		for($i=0; $i<$cnt-1; $i++){
			for ($newCol = $col; $newCol <= $this->lastCol; ++$newCol) {
				$cellValue = $sheet->getCellByColumnAndRow($newCol, $row)->getValue();

				if (isset($cellValue)) {

					logit(" set cell value: " . $cellValue);
					$sheet->setCellValueByColumnAndRow($newCol, $newRow, $cellValue);
				}
			}
			$newRow++;
		}
	}

	//生成二维码并设置到对应的位置
	protected function setImg($imgPath, $col, $row, $width=75, $height=75) {
		logit("create img: " . $imgPath . " width: " . $width . " height: " . $height);

		$sheet = $this->sheet;
		//$drawing->setName('qrCode');

		//支持png和jpeg两种图片
		if (stripos($imgPath, "http") !== false) {
			$type = "jpg";
			$img = @imagecreatefromjpeg($imgPath);
			if($img == null || !$img) {
				logit("failed.....");

				$type = "png";
				$img = @imagecreatefrompng($imgPath);
				if($img == null || !$img) {
					return;
				}
				else {
					$renderType = PHPExcel_Worksheet_MemoryDrawing::RENDERING_PNG;
					$mimeType = PHPExcel_Worksheet_MemoryDrawing::MIMETYPE_PNG;
				}
			}
			else {
				$renderType = PHPExcel_Worksheet_MemoryDrawing::RENDERING_JPEG;
				$mimeType = PHPExcel_Worksheet_MemoryDrawing::MIMETYPE_JPEG;
			}
			//$width = imagesx($img);
			//$height = imagesy($img);

			$drawing =new PHPExcel_Worksheet_MemoryDrawing();
			$drawing->setImageResource($img);
			$drawing->setRenderingFunction($renderType);//渲染方法
			$drawing->setMimeType($mimeType);
		}
		else {
			$drawing = new PHPExcel_Worksheet_Drawing();
			$drawing->setPath($imgPath);
		}

		logit("image created: " . $imgPath);

		//col起始位置是A, 就是65
		$colName = chr(65 + $col) . $row;

		$drawing->setWidth($width);
		$drawing->setHeight($height);
		$drawing->setCoordinates($colName);
		$drawing->setOffsetX(15);
		$drawing->setOffsetY(2);
		$drawing->setWorksheet($sheet);

		$sheet->setCellValueByColumnAndRow($col, $row, "");
	}
}

function printFile($fmt, $ret, $fname, $tpl)
{
	if ($fmt != "excel") {
		return;
	}

	header("Content-disposition: attachment; filename=" . $fname . ".xlsx");
	header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
	header("Content-Transfer-Encoding: binary");

	$retArr = table2objarr($ret);

	//require_once 'class/vendor/autoload.php';
	require_once 'class/vendor/phpoffice/phpexcel/Classes/PHPExcel.php';
	require_once 'class/vendor/phpoffice/phpexcel/Classes/PHPExcel/IOFactory.php';
	require_once 'class/phpqrcode/phpqrcode.php';

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
		logit($retArr[0]);
		$print = new PrintUtil($sheet, $retArr[0], $highestRow, $highestColumnIndex);
	}
	else {
		logit($retArr);
		$print = new PrintUtil($sheet, $retArr, $highestRow, $highestColumnIndex);
	}

	for ($row = 1; $row <= $highestRow; ++$row) {
		for ($col = 0; $col <= $highestColumnIndex; ++$col) {
			$print->handleCell($col, $row, $highestRow);
		}
	}

	$writer->save('php://output');
	return true;
}

function autoNo() {
	static $i = 0;
	return ++$i;
}

function concatField() {
	$r = PrintUtil::$curVal;

	$numargs = func_num_args();
    logit("Number of arguments: $numargs");

	$concatStr = "";
    $arg_list = func_get_args();
    for ($i = 0; $i < $numargs; $i++) {
		if ($i > 0) {
			$concatStr .= '&';
		}

		$valStr = 'return ($r["' . $arg_list[$i] . '"]);';
		$val = eval($valStr);

		//logit("valStr: " . $valStr . " val: " . $val);
		$concatStr .= $arg_list[$i] . '=' . $val;
    }

	return $concatStr;
}

function qrcode($str) {
	$errorCorrectionLevel = 'L';//容错级别
	$matrixPointSize = 6;//生成图片大小
	$imgPath = './qrcode-' . autoNoPrivate() . '.png';
	QRcode::png($str, $imgPath, $errorCorrectionLevel, $matrixPointSize, 2);

	logit("image str: " . $str . " image path: " . $imgPath);

	return $imgPath;
}

function autoNoPrivate() {
	static $noPrivate = 0;
	return ++$noPrivate;
}

function getField($v) {
	$r = PrintUtil::$curVal;

	$valStr = 'return ($r["' . $v . '"]);';
	logit("getField str: " . $valStr);

	return eval($valStr);
}

/*
function api_testMatch() {
	$src = '{a::b} + {b}';//mparam("src");

//	$expr = preg_replace_callback('/(?<![0-9])\$?([a-z_]\w*\b)(?!\()/iU', function ($ms) {
	$expr = preg_replace_callback('/\{([^{}]+)\}/iU', function ($ms) {
		$m = $ms[0];
	
		addLog("start");
		addLog($ms[0]);
		addLog($ms[1]);
		addLog("end");

		$expr = preg_replace_callback('/(?<![0-9])\$?([a-z_]\w*\b)(?!\()/iU', function ($ms2) {
			addLog("2start");
			addLog($ms2[0]);
			addLog($ms2[1]);
			addLog("2end");
			return ">2<";
		}, $ms[1]);

		addLog($expr);
		//logit("ms[0]: " . $m . " expr: " . $expr);
		return ">1<";
	}, $src);

	addLog($expr);
}
*/
// vi: foldmethod=marker