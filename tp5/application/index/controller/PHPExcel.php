<?php
namespace app\index\controller;
use think\Controller;
use app\index\controller\Base;
use app\index\controller\Location;
use app\index\model\Shop;
use think\Cache;
use think\Db;
use think\Url;
use think\Config;
class Phpexcel extends Base
{

function intoExcel(){  
   if (!empty($_FILES)){
       // import('PHPExcel.PHPExcel', EXTEND_PATH);
       vendor("PHPExcel.PHPExcel");
       // 导入PHPExcel类库
       $PHPExcel = new \PHPExcel();
       // 创建PHPExcel对象，注意，不能少了
        $file = request()->file('file1');
        $info = $file->validate(['size'=>1024*1024*1024,'ext'=>'xlsx,xls,csv'])->move(ROOT_PATH . 'Uploads' . DS . 'Excel');
        if($info){
        $exclePath = $info->getSaveName();
        //获取文件名
        $file_name = ROOT_PATH . 'Uploads' . DS . 'Excel' . DS . $exclePath;
        //上传文件的地址
        $objReader =\PHPExcel_IOFactory::createReader('Excel2007');
        $obj_PHPExcel =$objReader->load($file_name, $encode = 'utf-8');
        //加载文件内容,编码utf-8
        $excel_array=$obj_PHPExcel->getsheet(0)->toArray();
        //转换为数组格式
        var_dump($excel_array);
      } else{
      // 上传失败获取错误信息
        echo $file->getError();
      }
    }
}

/** 
 * 创建(导出)Excel数据表格 
 * @param  array   $list        要导出的数组格式的数据 
 * @param  string  $filename    导出的Excel表格数据表的文件名 
 * @param  array   $indexKey    $list数组中与Excel表格表头$header中每个项目对应的字段的名字(key值) 
 * @param  array   $startRow    第一条数据在Excel表格中起始行 
 * @param  [bool]  $excel2007   是否生成Excel2007(.xlsx)以上兼容的数据表 
 * 比如: $indexKey与$list数组对应关系如下: 
 *     $indexKey = array('id','username','sex','age'); 
 *     $list = array(array('id'=>1,'username'=>'YQJ','sex'=>'男','age'=>24)); 
 */  
function exportExcel($list,$filename,$indexKey,$fieldArray,$widthArray){  
        $xlsData = $list;
        Vendor('PHPExcel.PHPExcel');//调用类库,路径是基于vendor文件夹的
        Vendor('PHPExcel.PHPExcel.Worksheet.Drawing');
        Vendor('PHPExcel.PHPExcel.Writer.Excel2007');
        $objExcel = new \PHPExcel();
        //set document Property
        $objWriter = \PHPExcel_IOFactory::createWriter($objExcel, 'Excel2007');
 
        $objActSheet = $objExcel->getActiveSheet();
        $key = ord("A");
        $letter =explode(',',"A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z");
        $arrHeader = $indexKey;
        //填充表头信息
        $lenth =  count($arrHeader);
        for($i = 0;$i < $lenth;$i++) {
            $objActSheet->setCellValue("$letter[$i]1","$arrHeader[$i]");
        };
        //填充表格信息
        // 判断所传表格2018.7.5增加'昵称','姓名',
         
    

                foreach($xlsData as $k=>$v){
                    $k +=2;
                    for ($i=0; $i <count($fieldArray) ; $i++) { 
                       $objActSheet->setCellValue($letter[$i].$k,$v[$fieldArray[$i]]);
                    }
                    
                    // $objActSheet->setCellValue('A'.$k, $v['id']);
                    // // $objActSheet->setCellValue('B'.$k, $v['title']);
                    // // 图片生成
                    // $objDrawing[$k] = new \PHPExcel_Worksheet_Drawing();
                    // $objDrawing[$k]->setPath('public/static/admin/images/profile_small.jpg');
                    // // 设置宽度高度
                    // $objDrawing[$k]->setHeight(40);//照片高度
                    // $objDrawing[$k]->setWidth(40); //照片宽度
                    // /*设置图片要插入的单元格*/
                    // $objDrawing[$k]->setCoordinates('C'.$k);
                    // // 图片偏移距离
                    // $objDrawing[$k]->setOffsetX(30);
                    // $objDrawing[$k]->setOffsetY(12);
                    // $objDrawing[$k]->setWorksheet($objPHPExcel->getActiveSheet());
                    // 表格内容
                   
                    // 表格高度
                    $objActSheet->getRowDimension($k)->setRowHeight(20);
                }

         

        $width = array(5,10,15,20,25,30,35,40,45,50);
        //设置表格的宽度
        for ($i=0; $i <count($fieldArray) ; $i++) { 
         $objActSheet->getColumnDimension($letter[$i])->setWidth($width[$widthArray[$i]]);
        }
        // $objActSheet->getColumnDimension('A')->setWidth($width[1]);
        // $objActSheet->getColumnDimension('B')->setWidth($width[2]);
      
        
 
 
        $outfile = $filename.".xls";
        ob_end_clean();
        header("Content-Type: application/force-download");
        header("Content-Type: application/octet-stream");
        header("Content-Type: application/download");
        header('Content-Disposition:inline;filename="'.$outfile.'"');
        header("Content-Transfer-Encoding: binary");
        header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
        header("Pragma: no-cache");
        $objWriter->save('php://output');
}


      
}
