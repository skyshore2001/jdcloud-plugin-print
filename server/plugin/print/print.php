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
				$this->obj = isset($subObj) ? $this->ret[$subObj] : $ret;
				$this->startRow = $row;	
				$this->endRow = $row + $cnt;

				logit("start add row");
				//j-for这一个括号里的内容去掉
				$tempVal = preg_replace_callback('/\{([^{}]+)\}/iU', function ($ms2) {
					return "";
				}, $this->cellValue, 1);

				logit("temp val: " . $tempVal);
				$this->sheet->setCellValueByColumnAndRow($col, $row, $tempVal);
				$this->addRow($cnt, $row, $col);	
				$lastRow = $this->endRow;

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
					}
				}
				else {
					$r = $this->ret;
				}

				logit("before eval: " . $expr);

				$expr = "return (" . $expr . ");";

				logit("full expr: " . $expr);

				$ret = eval($expr);
				if (!isset($ret)) {
					$ret = "";
				}
				logit("after eval: " . $ret);
			}

			return $ret;
		}, $cellValue);

		if ($matched) {
			if ($isImg) {
				$this->setImg($expr, $col, $row); 
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
	protected function setImg($imgPath, $col, $row) {
		logit("create img: " . $imgPath);

		$sheet = $this->sheet;
		//$drawing->setName('qrCode');

		if (stripos($imgPath, "http") !== false) {
			$img = @imagecreatefromjpeg($imgPath);
			if($img == null || !$img) {
				logit("failed.....");
				return;
			}
		
			$drawing =new PHPExcel_Worksheet_MemoryDrawing();
			$drawing->setImageResource($img);
			$drawing->setRenderingFunction(PHPExcel_Worksheet_MemoryDrawing::RENDERING_JPEG);//渲染方法
			$drawing->setMimeType(PHPExcel_Worksheet_MemoryDrawing::MIMETYPE_JPEG);
		}
		else {
			$drawing = new PHPExcel_Worksheet_Drawing();
			$drawing->setPath($imgPath);
		}

		//col起始位置是A, 就是65
		$colName = chr(65 + $col) . $row;

		$drawing->setHeight(74);
		$drawing->setWidth(74);
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

	require_once 'class/vendor/autoload.php';
	require_once 'class/vendor/phpoffice/phpexcel/Classes/PHPExcel.php';
	require_once 'class/vendor/phpoffice/phpexcel/Classes/PHPExcel/IOFactory.php';
	require_once 'class/phpqrcode/phpqrcode.php';

	//require_once "D:/work/wms/server/php/class/vendor/autoload.php";
	//require_once "D:/work/wms/server/php/class/vendor/phpoffice/phpexcel/Classes/PHPExcel/Autoloader.php";
	//require_once "D:/work/wms/server/php/class/vendor/phpoffice/phpexcel/Classes/PHPExcel.php";
	//require_once "D:/work/wms/server/php/class/vendor/phpoffice/phpexcel/Classes/PHPExcel/IOFactory.php";
	//require_once "D:/work/wms/server/php/class/phpqrcode/phpqrcode.php";

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
	$numargs = func_num_args();
    logit("Number of arguments: $numargs");

	$concatStr = "";
    $arg_list = func_get_args();
    for ($i = 0; $i < $numargs; $i++) {
		if ($i > 0) {
			$concatStr .= '&';
		}
		$concatStr .= $arg_list[$i] . '=$r[' . $arg_list[$i] . ']';
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