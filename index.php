<?php
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
// Проверка подключена ли библиотека
if (file_exists(__DIR__ . '/vendor') && file_exists(__DIR__ . '/vendor/phpoffice/phpspreadsheet')) {
    require 'vendor/autoload.php';
    //Проверка существует ли директория для записи файла
    if (file_exists(__DIR__ . '/file')) {
        // Проверка доступен ли файл для записи
        if (@fopen(__DIR__ . '/file/random_numbers.xlsx', 'w+')) {
            // Создание массива для заполнения таблицы
            $numbers = array();
            for ($i = 0; $i < 10; $i++) {
                $numbers[$i] = array();
                for ($j=0; $j < 10; $j++) {
                    $numbers[$i][$j] = mt_rand(1, 100);
                }
            }
            //Создание и заполнение таблицы
            $spreadsheet = new Spreadsheet();
            $spreadsheet->getActiveSheet()
                ->fromArray(
                    $numbers,
                    NULL,
                    'A1'
                );
            //Установка стилей для таблицы
            $styleArray = [
                'alignment' => [
                    'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                    'verical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                ],
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    ],
                ]
            ];
            $columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'];
            $spreadsheet->getActiveSheet()->getStyle('A1:J10')
                ->applyFromArray($styleArray);
            foreach ($columns as $column) {
                $spreadsheet->getActiveSheet()->getColumnDimension($column)
                ->setWidth(60, 'px');
            }
            //Создание и сохранение нового Excel файла
            $writer = new Xlsx($spreadsheet);
            $writer->save(__DIR__ . '/file/random_numbers.xlsx');
            echo 'Файл создан';
        } else {
            echo 'файл открыт. Закройте перед созданием нового файла';
        }
    } else {
        echo 'На сервере отсутствует директория file';
    }
} else {
    echo 'Библиотека PhpSpreadsheet не подключена';
}



