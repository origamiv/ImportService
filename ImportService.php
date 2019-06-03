<?php

namespace App\Services;

use PDO;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Reader;
use Throwable;

/**
 * Class ImportService  - сервис для импорта файлов в любую таблицу БД из Excel файлов
 * Пример использования - в модуле timetable
 * @author Veniamin Smorodinsky *
 * @package app\services
 */
class ImportService
{
    public $db;
    public $filename;       //имя импортируемого файла
    public $spreadsheet;
    public $worksheet;
    public $captions=[];       // соответствие заголовков первой строк и колонок таблицы
    public $data=[];           // данные из таблицы
    public $fields = [];         // соответствие полей и заголовков
    public $columns=[];        // соответствие колонок и названий полей
    public $rows = [];        // преобразованные данные, готовые для вставки в таблицу
    public $tables_fields=[];  // соответствие таблиц и полей [fld => table]
    public $tables_descr=[];   // соответствие таблиц и полей
    //  ['table' => [
    //     'fields'=>['field1', 'field2', 'field3'],
    //     'default'=>['fld1'=>'val1']
    //  ]
    public $rules;          // правила вычислений и преобразований для полей
    public $tables_data;
    public $uniq = [];        // массив для сохранения уникальных значений при функции uniq_sort

    public function __construct(PDO $db)
    {
        $this->db=$db;
    }

    public function read($fn)
    {
        $this->filename = $_SERVER['DOCUMENT_ROOT'] . DIRECTORY_SEPARATOR . $fn;
       
        $reader = new Reader\Xlsx();
        $this->spreadsheet = $reader->load($this->filename);
        $this->worksheet = $this->spreadsheet->getActiveSheet();


        $captions = [];
        $data = [];
        foreach ($this->worksheet->getRowIterator() as $rowIndex => $row) {
            $cellIterator = $row->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(FALSE);
            foreach ($cellIterator as $k => $cell) {
                $val = $cell->getFormattedValue(); //Value();
                $address = $cell->getCoordinate();

                if ($rowIndex == 1) {
                    $captions[$k] = $val;
                } else {
                    $data[$rowIndex][$k] = $val;
                }
            }
        }



        $this->captions = $captions;
        $this->data = $data;

    }

    public function convert($fld, $val, $item = [])
    {

        if (isset($this->rules[$fld])) {
            $rule = $this->rules[$fld];

            if ($rule['fld'] == $fld) {
                $first_val = $val;
            } else {
                $fld_name = $rule['fld'];
                $first_val = $item[$fld_name];
                $val = $first_val;
            }

            if (isset($rule['func'])) {

                if ($rule['func'] == 'to_lower') {
                    $val = strtolower($first_val);
                }

                if ($rule['func'] == 'to_time') {
                    $val = strtotime($first_val);
                }

                if ($rule['func'] == 'uniq_sort') {
                    if (isset($this->uniq[$fld][$first_val]) == false) {
                        $this->uniq[$fld][$first_val] = 1 + count($this->uniq[$fld]);
                        $val = $this->uniq[$fld][$first_val];
                    } else {
                        $val = $this->uniq[$fld][$first_val];
                    }
                }

                if ($rule['func'] == 'concat') {
                    $params = explode(',', $rule['params']);
                    $val2 = '';
                    foreach ($params as $param) {
                        $fld_name = trim($param);
                        $val2 = $val2 . ' ' . $item[$fld_name];
                    }

                    $val = $val2;
                }

                if ($rule['func'] == 'get_id_from_link_table') {
                    $params = explode(',', $rule['params']);
                    $link_table = $params[0];
                    $link_id = $params[1];
                    $link_val = $params[2];

                    $sql = "SELECT $link_id as id FROM $link_table WHERE $link_val='$first_val'";
                    $r = Yii::$app->db->createCommand($sql)->queryOne();
                    $val = $r['id'];
                }

                if ($rule['func'] == 'date_format') {
                    $val = date($rule['params'], strtotime($first_val));
                }

            }

        }

        return $val;
    }

    public function setFields($tpl, $tables_fields, $rules)
    {

        //echo "asda";
        //print_r($this->data);
        //exit;

        $this->rules = $rules;
        $this->tables_descr = $tables_fields;
        $this->tables_fields = [];

        foreach ($tables_fields as $table => $item) {
            foreach ($item['fields'] as $fld) {
                $this->tables_fields[$fld] = $table;
            }
        }

        $fld = [];
        $captions = [];

        foreach ($this->captions as $cell => $title) {
            $title = trim($title);
            $captions[$title] = $cell;
        }

        foreach ($tpl['excel'] as $title => $field) {
            $cell = $captions[$title];
            $fld[$field]['column'] = $cell;
            $fld[$field]['title'] = trim($title);

            $columns[$cell] = $field;
        }

        foreach ($tpl['computed'] as $title => $field) {
            $fld[$field]['title'] = trim($title);

            $columns_computed[] = $field;
        }

        $this->fields = $fld;
        $this->columns = $columns;

        $rows = [];

        // убираем вторую строку - с примечаниями
//
        unset($this->data[2]);

        //print_r($columns);

        foreach ($this->data as $data_row) {
            unset($row);

            foreach ($data_row as $col => $val) {
                if (isset($columns[$col])) {
                    $col_field = $columns[$col];

                    if ((trim($col_field) != '') && (trim($val) != '')) {
                        $row[$col_field] = $val;
                    }
                }
            }


            if (isset($row) && count($row) > 0) {
                if (isset($columns_computed) && (count($columns_computed)>0))
                {
                    foreach ($columns_computed as $col)
                    {
                        $row[$col] = '';
                    }
                }

                $rows[] = $row;
            }

            //print_r($row);

        }

        //print_r($rows);
        $this->rows = $rows;
    }

    public function compute()
    {
        $this->tables_data = [];

        foreach ($this->rows as $k => $item) {
            unset($data_item);
            foreach ($item as $fld => $val) {
                $table = $this->tables_fields[$fld];
                $item[$fld] = $this->convert($fld, $val, $item);
                $data_item[$table][$fld] = $item[$fld]; // $this->convert($fld, $val, $item);
            }


            foreach ($data_item as $table => $val_item) {
                $this->tables_data[$table][] = $val_item;
            }
        }
    }

    public function super_unique($array)
    {
        $result = array_map("unserialize", array_unique(array_map("serialize", $array)));

        foreach ($result as $key => $value) {
            if (is_array($value)) {
                $result[$key] = $this->super_unique($value);
            }
        }

        return $result;
    }

    public function run()
    {
        $result = [];
        $this->compute();

        //print_r($this->tables_data);

        foreach ($this->tables_data as $table => $tdata) {
            try {

                if (isset($this->tables_descr[$table]['before_load'])) {
                    foreach ($this->tables_descr[$table]['before_load'] as $command) {
                        if ($command == 'distinct') {
                            $this->tables_data[$table] = $this->super_unique($this->tables_data[$table]);
                        }
                    }

                }

                $fld_str = '';
                $val_str = '';
                foreach ($this->tables_descr[$table]['default'] as $fld => $value) {
                    if (trim($fld_str) != '') {
                        $fld_str .= ',';
                    }
                    $fld_str = $fld_str . '"'.$fld.'"';

                    if (trim($val_str) != '') {
                        $val_str .= ',';
                    }
                    $val_str = $val_str . "'$value'";
                }


                foreach ($this->tables_data[$table] as $k => $row)
                {
                    $dat = date('Y-m-d H:i:s');
                    $sql = 'INSERT INTO public.' . $table . ' (' . $fld_str . ') VALUES (' . $val_str . ') RETURNING id';
                    //echo $sql."\r\n";
                    $id=$this->db->query($sql)->fetch()['id'];
                    //print_r($r);
                    //$id = $this->db->lastInsertId();

                    //echo $id;
                    //print_r($this->tables_data[$table]);

                    $this->tables_data[$table][$k]['inserted_id'] = $id;

                    foreach ($row as $fld => $val) {
                        try {
                            $sql = "UPDATE public.$table SET \"$fld\"='$val' WHERE id=$id";
                            //echo $sql."\r\n";
                            $this->db->query($sql)->execute();
                        } catch (Throwable $e) {

                            $result['success'] = false;
                            $result['error']['fld'] = $fld;
                        }
                    }



                    if (isset($this->tables_descr[$table]['after_insert'])) {
                        if ($this->tables_descr[$table]['after_insert']['func'] == 'set_link_id') {
                            $link_fld_nam = $this->tables_descr[$table]['after_insert']['params'];
                            $link_table = $this->tables_fields[$link_fld_nam];
                            $this->tables_data[$link_table][$k][$link_fld_nam] = $id;
                        }

                        if ($this->tables_descr[$table]['after_insert']['func'] == 'set_link_id_from_data') {

                            $params = explode(',', $this->tables_descr[$table]['after_insert']['params']);


                            $link_table = $params[0];
                            $link_fld_compare = $params[2];

                            $curr_val=$row[$link_fld_compare];

                            $link_fld_from = $this->tables_descr[$table]['after_insert']['fld_from'];
                            $link_fld_to   = $this->tables_descr[$table]['after_insert']['fld_to'];
                            $linked_table  = $this->tables_fields[$link_fld_from];

                            foreach ($this->tables_data[$linked_table] as $nom_row => $item2)
                            {
                                foreach ($this->tables_data[$link_table] as $linked_item)
                                {
                                    if (($item2[$link_fld_from] == $linked_item[$link_fld_compare]) && ($item2[$link_fld_from] == $curr_val))
                                    {
                                        $this->tables_data[$linked_table][$nom_row][$link_fld_to] = $linked_item['inserted_id'];
                                    }
                                }
                            }
                        }
                    }
                }

                $result['success'] = true;
            } catch (Throwable $e) {
                $result['success'] = false;
                $result['error']['table'] = $table;
            }
        }

        return $result;
    }

}
