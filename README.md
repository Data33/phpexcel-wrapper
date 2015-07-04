# phpexcel-wrapper
A quick and easy wrapper to make excel exports easier

## Example usage

```php
$excel = new Data33\ExcelWrapper\ExcelWrapper();

$excel->setTitle('My first excel file')
		->addRow(['Country', 'Capital'], 'header')
		->addRow(['Sweden', 'Stockholm'])
		->addRow(['Norway', 'Oslo'])
		->save('countries.xlsx');
```

To style specific cells:

```php
$excel->setTitle('My first excel file')
		->addRow(['Country', 'Capital'], 'header')
		->addRow([['Europe', 'header']])
		->addRow(['Sweden', 'Stockholm'])
		->addRow([['Africa', 'header']])
		->addRow(['Tunisia', 'Tunis'])
		->save('countries.xlsx');
```

To add custom styles we can give the wrapper PHPExcel style arrays:

```php
Data33\ExcelWrapper\ExcelStyle::setStyle('red', [
	'font' => [
		'size' => 10,
		'name' => 'Arial',
		'color' => [
			'rgb' => 'ff0000'
		]
	]
]);

$excel->setTitle('My first excel file')
		->addRow(['Country', 'Capital'], 'header')
		->addRow(['Sweden', ['Stockholm', 'red']])
		->addRow(['Norway', ['Oslo', 'red']])
		->save('countries.xlsx');
```

To output directly to browser for download:

```php
$excel->setTitle('My first excel file')
		->addRow(['Country', 'Capital'], 'header')
		->addRow(['Sweden', 'Stockholm'])
		->addRow(['Norway', 'Oslo'])
		->outputToBrowser('countries.xlsx');
```