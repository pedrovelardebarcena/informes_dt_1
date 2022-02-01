<?php

namespace App\Http\Controllers;
aa
use PDOException;
use Exception;
use Illuminate\Database\QueryException;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Facades\Log;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Pdf\Mpdf;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

class informe1 extends Controller
{
    
    function get_categorias() {

        $sql = "SELECT A.id, A.codigo, A.descripcion, A.modtimestamp 
            FROM CODIGOS A WHERE A.concepto = 4
            ORDER BY A.descripcion";

        return DB::select( $sql );

    }    

    public function generar() {

        function generar_cabecera($sheet, $fila) {

            $sheet->setCellValue("B" . $fila, "#");
            $sheet->setCellValue("C" . $fila, "Código");
            $sheet->setCellValue("D" . $fila, "Descripcion");

            $rango = $sheet->getStyle("B" . $fila . ":D" . $fila);
            $rango->getFont()->setBold(true);
            $rango->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB("d9d9d9");

        }

        function generar_bordes($sheet, $fila_inicial, $fila_final) {

            $rango = $sheet->getStyle("B" . $fila_inicial . ":D" . $fila_final);

            $rango->getBorders()
                ->getAllBorders()
                ->setBorderStyle(Border::BORDER_THIN);

            $rango->getAlignment()->setHorizontal("left");

        }


        $fichero_nombre = "categorias";

        $spreadsheet = new Spreadsheet();

        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle("Categorias");

        $sheet->getColumnDimension('A')->setWidth(4);
        $sheet->getColumnDimension('B')->setWidth(6);
        $sheet->getColumnDimension('C')->setWidth(15);
        $sheet->getColumnDimension('D')->setWidth(80);

        $sheet->setCellValue("B2", "LABOROFFICE");
        $sheet->setCellValue("B3", "Categorias");
        $sheet->setCellValue("B4", date("d/m/Y H:i:s"));
        $sheet->getStyle("B2:B5")->getFont()->setBold(true);

        $dataset = $this->get_categorias();

        $fila = 6;

        generar_cabecera($sheet, $fila);

        foreach ($dataset as $item) {

            $fila = $fila + 1;

            $sheet->setCellValue("B" . $fila, $item->id);
            $sheet->setCellValue("C" . $fila, $item->codigo);
            $sheet->setCellValue("D" . $fila, $item->descripcion);

        }        

        generar_bordes($sheet, 6, $fila);

        $sheet->getStyle("A1");
       

        // EXPORTAR EN FORMATO EXCEL
        $writer = new Xlsx($spreadsheet);
        $writer->save($fichero_nombre . ".xlsx");

        // header("Access-Control-Allow-Origin: *");
        // header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        // $writer->save('php://output');

        // EXPORTAR EN FORMATO PDF;
        // $writer = new Mpdf($spreadsheet);
        // $writer->writeAllSheets();
        // $writer->save($fichero_nombre . ".pdf");

        // header("Access-Control-Allow-Origin: *");
        // header('Content-Type: application/pdf');
        // $writer->save('php://output');

        return response()->download($fichero_nombre . ".xlsx")->deleteFileAfterSend(true);

    }

}
