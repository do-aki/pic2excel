<?php
require dirname(__FILE__) . '/FastPHPExcel.php';

if (!ini_get('date.timezone')) {
    date_default_timezone_set('Asia/Tokyo');
}

function x2cell($n) {
    $r = range('A', 'Z');
    return ($n < 26) ? $r[$n] : $r[$n%26] . x2cell(floor($n/26)-1);
}

function xy2cell($x, $y) {
    return strrev(x2cell($x)) . ($y+1);
}

function createImage($filename) {
    $ext2function = array(
        'gd2' => 'gd2',
        'gd' => 'gd',
        'gif' => 'gif',
        'jpeg' => 'jpeg',
        'jpg' => 'jpeg',
        'png' => 'png',
        'wbmp' => 'wbmp',
        'xbm' => 'xbm',
        'xpm' => 'xpm',
    );
    
    $ext = pathinfo($filename, PATHINFO_EXTENSION);
    isset($ext2function[$ext]) or die('not support image');
    return call_user_func("imagecreatefrom{$ext2function[$ext]}", $filename);
}

function pic2excel($filename) {
    $excel = new FastStylePHPExcel();
    $excel->getProperties()->setCreator('do_aki')
        ->setLastModifiedBy('do_aki');

    $output_file = pathinfo($filename, PATHINFO_FILENAME) . '.xlsx';
    $img = createImage($filename);

    $img_x = imagesx($img);
    $img_y = imagesy($img);

    $color = new PHPExcel_Style_Color();
    print "building... ";
    
    $sheet = $excel->getActiveSheet();
    for ($y=0; $y < $img_y; ++$y) {
        for ($x=0; $x < $img_x; ++$x) {
            printf("\rbuilding... (%3d, %3d)   ", $x, $y);
            $rgb = imagecolorat($img, $x, $y);
            $rgb = sprintf("%02X%02X%02X", ($rgb >> 16) & 0xFF, ($rgb >> 8) & 0xFF, $rgb & 0xFF);
            $color->setRGB($rgb);

            $sheet->getStyle(xy2cell($x, $y))->getFill()
                ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                ->setStartColor($color);
        }
    }

    print "\nresizing...";
    for($i=0;$i<$img_x;++$i) {
        $sheet->getColumnDimensionByColumn($i)->setWidth(0.3);
    }
    for($i=0;$i<=$img_y;++$i) {
        $sheet->getRowDimension($i)->setRowHeight(2);
    }

    PHPExcel_IOFactory::createWriter($excel, 'Excel2007')->save($output_file);
    print "\nsaved: {$output_file}";
}

if ($argc <= 1 || !file_exists($argv[1])) {
    print "usage: {$argv[0]} image_file";
}

pic2excel($argv[1]);
