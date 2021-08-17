

## About Project

This project is to convert a raw data of expenses of customers to more organized way. Row data sample is located in storage/app/public folder named *document.xlsx*.

## Checking the project

To run and test the results you need to execute some simple commands:
 - composer install
 - php artisan excel:parse

This script will get the file automatically, format it and generate a new one in the same folder, named result-[datetime].xlsx
