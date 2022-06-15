<?php
Include('koneksi2.php');
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'no');
$sheet->setCellValue('B1', 'jenis');
$sheet->setCellValue('C1', 'masuk');
$sheet->setCellValue('D1', 'nis'); 
$sheet->setCellValue('E1', 'ujian');
$sheet->setCellValue('F1', 'paud');
$sheet->setCellValue('G1', 'tk');
$sheet->setCellValue('H1', 'skhun');
$sheet->setCellValue('I1', 'ijazah');
$sheet->setCellValue('J1', 'hobi');
$sheet->setCellValue('K1', 'cita');
$sheet->setCellValue('L1', 'nama');
$sheet->setCellValue('M1', 'kelamin');
$sheet->setCellValue('N1', 'nisn');
$sheet->setCellValue('O1', 'nik');
$sheet->setCellValue('P1', 'tempat');
$sheet->setCellValue('Q1', 'tgl');
$sheet->setCellValue('R1', 'agama');
$sheet->setCellValue('S1', 'kebutuhan');
$sheet->setCellValue('T1', 'alamat');
$sheet->setCellValue('U1', 'rt');
$sheet->setCellValue('V1', 'rw');
$sheet->setCellValue('W1', 'dusun');
$sheet->setCellValue('X1', 'kelurahan');
$sheet->setCellValue('Y1', 'kecamatan');
$sheet->setCellValue('Z1', 'pos');
$sheet->setCellValue('AA1', 'tinggal');
$sheet->setCellValue('AB1', 'transportasi');
$sheet->setCellValue('AC1', 'hp');
$sheet->setCellValue('AD1', 'telp');
$sheet->setCellValue('AE1', 'email');
$sheet->setCellValue('AF1', 'kip');
$sheet->setCellValue('AG1', 'nokip');
$sheet->setCellValue('AH1', 'warga');

$query = mysqli_query($koneksi, "select * from regis_peserta");
$i = 2; 
$no = 1;
while($row = mysqli_fetch_array($query))
{
	$sheet->setCellValue('A'.$i, $no++);
	$sheet->setCellValue('B'.$i, $row['jenis']);
	$sheet->setCellValue('C'.$i, $row['masuk']);
	$sheet->setCellValue('D'.$i, $row['nis']);
	$sheet->setCellValue('E'.$i, $row['ujian']);
	$sheet->setCellValue('F'.$i, $row['paud']);
	$sheet->setCellValue('G'.$i, $row['tk']);
	$sheet->setCellValue('H'.$i, $row['skhun']);
	$sheet->setCellValue('I'.$i, $row['ijazah']);
	$sheet->setCellValue('J'.$i, $row['hobi']);
	$sheet->setCellValue('K'.$i, $row['cita']);
	$sheet->setCellValue('L'.$i, $row['nama']);
	$sheet->setCellValue('M'.$i, $row['kelamin']);
	$sheet->setCellValue('N'.$i, $row['nisn']);
	$sheet->setCellValue('O'.$i, $row['nik']);
	$sheet->setCellValue('P'.$i, $row['tempat']);
	$sheet->setCellValue('Q'.$i, $row['tgl']);
	$sheet->setCellValue('R'.$i, $row['agama']);
	$sheet->setCellValue('S'.$i, $row['kebutuhan']);
	$sheet->setCellValue('T'.$i, $row['alamat']);
	$sheet->setCellValue('U'.$i, $row['rt']);
	$sheet->setCellValue('V'.$i, $row['rw']);
	$sheet->setCellValue('W'.$i, $row['dusun']);
	$sheet->setCellValue('X'.$i, $row['kelurahan']);
	$sheet->setCellValue('Y'.$i, $row['kecamatan']);
	$sheet->setCellValue('Z'.$i, $row['pos']);
	$sheet->setCellValue('AA'.$i, $row['tinggal']);
	$sheet->setCellValue('AB'.$i, $row['transportasi']);
	$sheet->setCellValue('AC'.$i, $row['hp']);
	$sheet->setCellValue('AD'.$i, $row['telp']);
	$sheet->setCellValue('AE'.$i, $row['email']);
	$sheet->setCellValue('AF'.$i, $row['kip']);
	$sheet->setCellValue('AG'.$i, $row['nokip']);
	$sheet->setCellValue('AH'.$i, $row['warga']);
	$i++;
}

$styleArray = [
	'borders' => [
		'allBorders' => [ 
			'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
		],
	],
];
$i = $i - 1;
$sheet->getStyle('A1:AH'.$i)->applyFromArray($styleArray);


$writer = new Xlsx($spreadsheet); 
$writer->save('Report Data PPDB.xlsx');
?>