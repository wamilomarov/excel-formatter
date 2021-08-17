<?php

namespace App\Console\Commands;

use Box\Spout\Common\Exception\InvalidArgumentException;
use Box\Spout\Common\Exception\IOException;
use Box\Spout\Common\Exception\UnsupportedTypeException;
use Box\Spout\Reader\Exception\ReaderNotOpenedException;
use Box\Spout\Writer\Common\Creator\Style\StyleBuilder;
use Box\Spout\Writer\Exception\WriterNotOpenedException;
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
     * @throws IOException
     * @throws InvalidArgumentException
     * @throws UnsupportedTypeException
     * @throws ReaderNotOpenedException
     * @throws WriterNotOpenedException
     */
    public function handle(): int
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
                    $data[$currentClient] = [
                        'Client' => $currentClient
                    ];
                }
            } else {
                // if no client yet selected
                if (!is_null($currentClient)) {
                    // if charge value already exists, increment it. otherwise, just create new one
                    $chargeExists = isset($data[$currentClient][$charge]);
                    if ($chargeExists) {
                        $data[$currentClient][$charge] += $value;
                    } else {
                        $data[$currentClient][$charge] = $value;
                    }
                }
            }
        }

        $headerStyle = (new StyleBuilder())->setFontBold()->setFontColor('0000ff')->build();

        $exportedFileName =Storage::disk('public')->path("result-" . date("Y-m-d-H-i-s") . ".xlsx");
        (new FastExcel($data))
            ->headerStyle($headerStyle)
            ->export($exportedFileName, function ($datum) {
            $result = [
                'Client' => $datum['Client']
            ];

            foreach (self::FIELDS as $field) {
                $result[$field] = $datum[$field] ?? 0;
            }
            return $result;
        });

        return 0;
    }
}
