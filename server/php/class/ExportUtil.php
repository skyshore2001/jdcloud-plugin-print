<?php

trait ExportUtil
{
  protected function onHandleExportFormat($fmt, $ret, $fname)
  {
    $tpl = param("tpl");
    if (isset($tpl)) {
      return printFile($fmt, $ret, $fname, $tpl);
    }
    else {
      return false;
    }
  }
}
