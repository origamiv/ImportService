
$import = new ImportService();
$import->read('imports/' . $fn);
$import->setFields([
    'excel' => [
        'Дата' => 'day_date',
        'Язык' => 'lang',
        'Время начала' => 'start',
        'Время конца' => 'end',
        'Заголовок события' => 'name',
        'Описание' => 'value',
        'Спикеры' => 'speaker_firstname',
        //'Зал'=>'',
        'Место проведения' => 'venue',
        'Номер дня' => 'day',
    ],
    //-------- полей ниже нет в Excel, но они должны вычисляться
    'computed' => [
        'Время начала unix' => 'start_at_ts',
        'Время конца unix' => 'end_at_ts',
        'Номер дня для таблицы дней' => 'day_id',
        'id спикера' => 'presentation_leader_id',
        'id event' => 'agdevent_id',
        'Номер дня' => 'day_nom',
    ]
],
    [
        'agd_days' => [
            'fields' => ['day', 'day_date'],
            'default' => ['status' => 'regular', 'ts' => $dat, 'room_id' => $roomId, 'day_date' => $dat],
            'before_load' => ['distinct'],
            //'after_insert'=>['func'=>'set_link_id','params'=>'day_id']
            'after_insert' => ['func' => 'set_link_id_from_data', 'params' => 'agd_days,inserted_id,day', 'fld_from' => 'day_nom', 'fld_to' => 'day_id']
        ],
        'agd_events' => [
            'fields' => ['day_nom', 'day_id', 'lang', 'start', 'end', 'start_at_ts', 'end_at_ts', 'name', 'value', 'speaker_firstname', 'venue'],
            'default' => ['status' => 'regular', 'ts' => $dat],
            'after_insert' => ['func' => 'set_link_id', 'params' => 'agdevent_id']
        ],
        'agdevents_presentation_leader' => [
            'fields' => ['agdevent_id', 'presentation_leader_id'],
            'default' => []
        ]
    ],
    [
        'lang' => ['fld' => 'lang', 'func' => 'to_lower'],
        'day_date' => ['fld' => 'day_date', 'func' => 'date_format', 'params' => 'Y-m-d'],
        //'day'=>['fld'=>'day_date', 'func'=>'uniq_sort'],
        'day_nom' => ['fld' => 'day'],
        //'day_id'=>['fld'=>'day_id'],
        //'day_id'=>['fld'=>'day_nom', 'func'=>'get_id_from_link_data', 'params'=>'agd_days,inserted_id,day'],
        'start' => ['fld' => 'start', 'func' => 'concat', 'params' => 'day_date, start'],
        'end' => ['fld' => 'end', 'func' => 'concat', 'params' => 'day_date, end'],
        'start_at_ts' => ['fld' => 'start', 'func' => 'to_time'],
        'end_at_ts' => ['fld' => 'end', 'func' => 'to_time'],
        'presentation_leader_id' => ['fld' => 'speaker_firstname', 'func' => 'get_id_from_link_table', 'params' => 'presentation_leader,id,fio']
    ]
);

$status = $import->run();
return $status;