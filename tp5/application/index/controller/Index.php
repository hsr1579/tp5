<?php
namespace app\index\controller;

use think\Controller;

use app\index\controller\PHPExcel;

use think\Db;
use think\Url;
use think\Config;
use think\Image;
class Index extends Controller
{
  
    public function daoru()
    {
      set_time_limit(0);
      
        return $this->fetch();
       }

    // 导入Excel表格
    public function intoExcel(){
       if (!empty($_FILES)){
           // import('PHPExcel.PHPExcel', EXTEND_PATH);
           vendor("PHPExcel.PHPExcel");
           // 导入PHPExcel类库
           $PHPExcel = new \PHPExcel();
           // 创建PHPExcel对象，注意，不能少了
            $file = request()->file('file1');
            $info = $file->validate(['size'=>10000000000,'ext'=>'xlsx,xls,csv'])->move(ROOT_PATH . 'Uploads' . DS . 'Excel');
            if($info){
            $exclePath = $info->getSaveName();
            //获取文件名
            $file_name = ROOT_PATH . 'Uploads' . DS . 'Excel' . DS . $exclePath;
            if (!file_exists($file_name)) {
                die('no file!');
            }
            $extension = strtolower( pathinfo($file_name, PATHINFO_EXTENSION) );
            
            if ($extension =='xlsx') {
               $objReader =\PHPExcel_IOFactory::createReader('Excel2007');
                //$objExcel = $objReader ->load($file);
            } else if ($extension =='xls') {
                $objReader =\PHPExcel_IOFactory::createReader('Excel5');
                //$objExcel = $objReader ->load($file);
            } else if ($extension=='csv') {
                $objReader =\PHPExcel_IOFactory::createReader('CVS');
            }

            //上传文件的地址
            
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
    // 到出Excel表格
    public function exportExcel(){
      // 表格数据
            $aa=db('lists')->where('nid',2)->select();
      // 表格名称
            $bb='测试总表';
      // 表头信息
          $cc=[
            'ID',
            '标题',
            '内容'
           ];
      // 数据对应字段
          $dd=[
            'id',
            'title',
            'content'
          ];
      // 表格宽度 1~10
          $ee=[
            '1',
            '1',
            '3'
          ];
            $Phpexcel=new Phpexcel;
            $Phpexcel->exportExcel($aa,$bb,$cc,$dd,$ee);
    }
   
}