<?php

namespace App\Console\Commands;

use App\Service\LoadService;
use Illuminate\Console\Command;
use Storage;

class Test extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = '123';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Command description';

    /**
     * Create a new command instance.
     *
     * @return void
     */
    public function __construct(LoadService $library_import_service)
    {
        parent::__construct();

        $this->library_import_service = $library_import_service;
    }

    /**
     * Execute the console command.
     *
     * @return mixed
     */
    public function handle()
    {
        $files = Storage::disk('read')->files('/');
        $xlsx_files = array_values(preg_grep('/.xlsx/m', $files));

        $xlsx_file = $xlsx_files[0];
        $xlsx_file_path = Storage::disk('read')->path($xlsx_file);

        $total_spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($xlsx_file_path);
        $total_excel_data = $this->library_import_service->getSpreadSheetData($total_spreadsheet);
        $total_excel_data = $total_excel_data[0];
        unset($total_excel_data[0]);

        $meta_template = Storage::path('excel_template/metadata.xlsx');
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($meta_template);

        $this->listAlbumSheet($spreadsheet, $total_excel_data);
        $this->listTrackSheet($spreadsheet, $total_excel_data);
        $this->listLabelCopySheet($spreadsheet, $total_excel_data);

        $new_meta_full_path = Storage::disk('out')->path('haha.xlsx');
        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
        $writer->save($new_meta_full_path);

        $spreadsheet->disconnectWorksheets();
        unset($spreadsheet);
    }

    private function listAlbumSheet($spreadsheet, $total_excel_data)
    {
        $total_excel_data = collect($total_excel_data);
        $insert_data = [];

        $filtered = $total_excel_data->filter(function ($value, $key) {
            return $value[15];
        });

        $grouped_by_album_name = $filtered->whereNotNull(15)->groupBy(6);

        foreach ($grouped_by_album_name as $group) {
            $upc = '';
            $album_name = '';
            $total_artist_array = [];
            foreach ($group as $row) {
                if ($row[16]) {
                    $upc = $row[16];
                }
                if ($row[5]) {
                    $album_name = $row[5];
                }
                $artirst_string = $row[8];
                $artist_array = preg_split("/,|&/m", $artirst_string);
                $total_artist_array = array_merge($total_artist_array, $artist_array);
            }

            foreach ($total_artist_array as $index => $item) {
                $value = trim($item);
                $pattern = '/Feat.*/m';
                $value = preg_replace($pattern, '', $value);
                $total_artist_array[$index] = $value;
            }

            $total_artist_array = array_unique($total_artist_array);
            $role_array = [];
            foreach ($total_artist_array as $index => $item) {
                $role_array[] = '主唱';
            }

            $total_artist_string = implode("/", $total_artist_array);
            $role_string = implode("/", $role_array);

            $insert = [
                $upc, $album_name, null, $total_artist_string, $role_string, null, null, count($group), '嘻哈與饒舌', null, '英文', null, null, '否',
            ];

            $insert_data[] = $insert;
        }

        $spreadsheet
            ->getSheet(0)
            ->insertNewRowBefore(5, count($insert_data))
            ->fromArray(
                $insert_data,
                'all_set',
                'A3'
            );
    }

    private function listTrackSheet($spreadsheet, $total_excel_data)
    {
        $total_excel_data = collect($total_excel_data);
        $insert_data = [];

        foreach ($total_excel_data as $row) {
            $isrc = $row[15];
            if (!$isrc) {
                continue;
            }
            $number = $row[4];
            $title = $row[7];
            $title = trim(preg_replace('/\[[^\]]*\]/m', '', $title));

            $artirst_string = $row[8];
            $artist_array = preg_split("/,|&/m", $artirst_string);

            $total_artist_array = [];
            $total_role_array = [];
            foreach ($artist_array as $index => $item) {
                $value = trim($item);
                $artist_array = explode("Feat.", $value);

                foreach ($artist_array as $index => $item) {
                    $value = trim($item);
                    $total_artist_array[] = $value;
                    if ($index == 0) {
                        $total_role_array[] = '主唱';
                    }

                    if ($index != 0) {
                        $total_role_array[] = '合唱(feat.)';
                    }
                }
            }

            $role_count = count($total_role_array);
            foreach (['作詞者', '作曲者'] as $role) {
                for ($index = 0; $index < $role_count; $index++) {
                    $total_role_array[] = $role;
                }
            }

            $total_artist_array = collect([$total_artist_array, $total_artist_array, $total_artist_array]);
            $total_artist_array = $total_artist_array->flatten()->all();
            $total_artist_string = implode("/", $total_artist_array);

            $role_string = implode("/", $total_role_array);

            $insert_data[] = [$isrc, null, 1, $number, $title, null, $total_artist_string, $role_string, '嘻哈與饒舌', null, '英文', null, '兒童不宜', '歐美男歌手', '流行音樂'];
        }

        $spreadsheet
            ->getSheet(1)
            ->insertNewRowBefore(5, count($insert_data))
            ->fromArray(
                $insert_data,
                'all_set',
                'A3'
            );
    }

    private function listLabelCopySheet($spreadsheet, $total_excel_data)
    {
        $total_excel_data = collect($total_excel_data);
        $insert_data = [];

        foreach ($total_excel_data as $row) {
            $isrc = $row[15];
            if (!$isrc) {
                continue;
            }

            $artirst_string = $row[8];
            $artist_array = preg_split("/,|&/m", $artirst_string);

            $total_artist_array = [];
            $total_role_array = [];
            foreach ($artist_array as $index => $item) {
                $value = trim($item);
                $artist_array = explode("Feat.", $value);

                foreach ($artist_array as $index => $item) {
                    $value = trim($item);
                    $total_artist_array[] = $value;
                }
            }

            $first_column = $isrc;
            foreach (['作曲者', '作詞者'] as $role) {

                $right = round(50 / count($total_artist_array), 2);

                foreach ($total_artist_array as $index => $data) {
                    if ($index == count($total_artist_array) - 1 && count($total_artist_array) > 1) {
                        $right = 50.00 - ($right * ($index));
                    }

                    $insert = [$first_column, $role, $data, null, 'FieryStar', $right, 'Y'];
                    $insert_data[] = $insert;

                    $first_column = null;
                }
            }
        }

        $spreadsheet
            ->getSheet(3)
            ->insertNewRowBefore(3, count($insert_data))
            ->fromArray(
                $insert_data,
                'all_set',
                'A2'
            );
    }
}
