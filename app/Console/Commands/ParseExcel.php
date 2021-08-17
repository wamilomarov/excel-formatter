<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use Illuminate\Contracts\Filesystem\FileNotFoundException;
use Illuminate\Support\Facades\Storage;
use Rap2hpoutre\FastExcel\FastExcel;

class ParseExcel extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'excel:parse';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Parses excel file to a new format';

    private $fileName = "document.xlsx";

    const FIELDS = [
        'FORMALITIES',
        'INSURANCE HANDLING FEE',
        'WORLD GATE',
        'CUSTOMS CLEARANCE'
    ];

    /**
     * Create a new command instance.
     *
     * @return void
     */
    public function __construct()
    {
        parent::__construct();
    }

    /**
     * Execute the console command.
     *
     * @return int
     * @throws FileNotFoundException
     */
    public function handle()
    {
        $path = Storage::disk('public')->path($this->fileName);
        $data = [];


        $lines = (new FastExcel())->import($path);

        $currentClient = null;
        foreach ($lines as $line) {
            $charge = trim($line['Charge']);
            $value = floatval(trim($line['Value']));
            // if not a charge that we need (so it is client)
            if (!in_array(mb_strtoupper($charge), self::FIELDS)) {
                // set it to current client
                $currentClient = $charge;
                // create a new client row with values if not exists
                if (!isset($data[$currentClient])) {
                    $data[$currentClient] = [];
                }
            } else {
                if (!is_null($currentClient)) {
                    $chargeExists = isset($data[$currentClient][$charge]);
                    if ($chargeExists) {
                        $data[$currentClient][$charge] += $value;
                    } else {
                        $data[$currentClient][$charge] = $value;
                    }
                }
            }
        }

        $headings = array_merge(['Client'], self::FIELDS);
        $print = [];
        foreach ($data as $client => $row) {
            $print[] = array_merge([$client], $row);
        }

        $this->table($headings, $print);

        return 0;
    }
}
