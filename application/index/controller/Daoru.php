<?php
namespace app\index\controller;

use think\Loader;
use think\Controller;

class Daoru extends Controller
{
    public function index() {
        return $this->fetch();
    }
    /**
     * 导入
     */
    public function do_excelImport() {
        $file = request()->file('file');
        $pathinfo = pathinfo($file->getInfo()['name']);
        $extension = $pathinfo['extension'];
        $savename = time().'.'.$extension;
        if($upload = $file->move('./upload',$savename)) {
            $savename = './upload/'.$upload->getSaveName();
            Loader::import('PHPExcel.PHPExcel');
            Loader::import('PHPExcel.PHPExcel.IOFactory.PHPExcel_IOFactory');
            $objReader = \PHPExcel_IOFactory::createReader('Excel2007');
            $objPHPExcel = $objReader->load($savename,$encode = 'utf8');
            $sheetCount = $objPHPExcel->getSheetCount();
            for($i=0 ; $i<$sheetCount ; $i++) {    //循环每一个sheet
                $sheet = $objPHPExcel->getSheet($i)->toArray();
                unset($sheet[0]);
                foreach ($sheet as $v) {


                    //字段依次累加，与数据库对应
                    $data['id'] = $v[0];
                    $data['username'] = $v[1];
                    $data['sex'] = $v[2];
                    $data['idcate'] = $v[3];
                    $data['dorm_id'] = $v[4];
                    $data['iclass'] = $v[5];
                    
                    try {
                        db('users1')->insert($data);
                    } catch(\Exception $e) {
                 
                         echo $e;exit;
                        return '插入失败';
                    }
                }
            }
            echo "succ";
        } else {
            return $upload->getError();
        }
    }
}
