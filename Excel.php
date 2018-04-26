<?php
/**
 * Created by PhpStorm.
 * User: 刘凯文
 * Date: 2018/4/26
 * Time: 18:27
 */
/*
*+----------------------------------------------------------------------
*   PHPExcel导出Excel表格
*   array $rearr  需要导出的数组
*+----------------------------------------------------------------------
*/
function export($rearr) {
    $result = array(
        '标题',
        '内容',
        '发表时间',
        '状态'
    );
    $arr = array(
        'A',
        'B',
        'C',
        'D'
    );
    //导入excel类
    include('PHPExcel.php');

    // 创建一个excel
    $objPHPExcel = new PHPExcel();

    /****************************************设置居中开始**************************************/
    foreach ($arr as $key => $value) {
        $objPHPExcel->getActiveSheet()->getStyle($value)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    }
    /****************************************设置居中结束**************************************/
    // 循环$arr定义的列设置每列内容居中
    $objPHPExcel->getProperties()->setCreator("Maarten Balliauw")->setLastModifiedBy("Maarten Balliauw")->setTitle("Office 2007 XLSX Test Document")->setSubject("Office 2007 XLSX Test Document")->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")->setKeywords("office 2007 openxml php")->setCategory("Test result file");

    /**************************************设置标题开始*****************************************/
    // 循环$arr定义的列和$result设置表头
    $objPHPExcel->setActiveSheetIndex(0);
    foreach ($arr as $key => $value) {
        $objPHPExcel->getActiveSheet()->setCellValue($value . "1", $result[$key]);
    }
    /**************************************设置标题结束*****************************************/

    /**************************************设置内容开始*****************************************/
    $objPHPExcel->setActiveSheetIndex(0);
    $i = 2;
    // $rearr需要导出的数据二维数组
    foreach ($rearr as $key => $value) {
        // 这里从二维数组里面通过键名获取到值放到相应的表格中
        $objPHPExcel->getActiveSheet()->setCellValue('A' . $i, $value['title']);
        $objPHPExcel->getActiveSheet()->setCellValue('B' . $i, $value['content']);
        $objPHPExcel->getActiveSheet()->setCellValue('C' . $i, $value['time']);
        $objPHPExcel->getActiveSheet()->setCellValue('D' . $i, $value['status']);
        $i++;
    }
    /**************************************设置内容结束*****************************************/

    /**************************************设置宽度开始*****************************************/
    // 循环$arr定义的列设置每列宽度
    foreach ($arr as $key => $value) {
        $objPHPExcel->getActiveSheet()->getColumnDimension($value)->setWidth(20);
    }
    /**************************************设置宽度结束*****************************************/

    /**************************************设置导出下载开始*****************************************/
    $objPHPExcel->getSheet(0)->setTitle('phpexcel'); // 工作区域标题
    header('Content-Type: application/vnd.ms-excel');
    header('Content-Disposition: attachment;filename="phpexcel测试.xls"');//导出文件
    header('Cache-Control: max-age=0');
    header('Cache-Control: max-age=1');
    header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
    header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
    header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
    header('Pragma: public'); // HTTP/1.0
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
    $objWriter->save('php://output');//导出文件
    /**************************************设置导出下载结束*****************************************/
}

//===========================================================================
//                      【注】：此处为重点                                  //
//===========================================================================

/*
 * $arr为一个数组(数据库中获取)
 *
 */
//示例
$arr = [
    0=>['title'=>'我是标题一','content'=>'我是内容一','time'=>'2018:4:26','status'=>'在线'],
    1=>['title'=>'我是标题二','content'=>'我是内容二','time'=>'2018:4:25','status'=>'离线']

];
export($arr);
