<?php

namespace App\Http\Controllers;

use App\Helpers\NewTemplateProcessor;
use App\Helpers\OpenTemplateProcessor;
use Illuminate\Http\Request;
use Maatwebsite\Excel\Facades\Excel;
use PhpOffice\PhpWord;

class HomeController extends Controller
{

	public function index()
	{
		return view('welcome');
	}

	public function generate(Request $request)
	{
		$this->validate($request, [
			'csv_file' => 'required|file|mimes:csv,txt',
			'logo' => 'required|image',
		]);
		$excel = Excel::load($request->file('csv_file'))->get();
		// $company = [];
		// foreach ($excel as $key => $data) {
		// 	$company[$data->qns_id] = $data->answer;
		// }
		
		// $phpWord = new PhpWord\PhpWord();
		// $section = $phpWord->addSection();
		// $headerStyle = ['alignment' => PhpWord\SimpleType\Jc::RIGHT];
		// $header = $section->addHeader();
		// $header->addImage($request->file('logo'), [
		// 	'positioning' => 'relative',
		// 	'height' => 40,
		// 	'wrappingStyle' => 'infront',
		// ]);

		// $header->addText(htmlspecialchars($company[2].' '.$company[3]), null, $headerStyle);
		// $lineStyle = ['weight' => 1.5, 'width' => 600, 'height' => 0, 'color' => 000000];
		// $header->addLine($lineStyle);

		// $footerStyle = ['alignment' => PhpWord\SimpleType\Jc::RIGHT];
		// $footer = $section->addFooter();
		// $lineStyle = ['weight' => 1, 'width' => 600, 'height' => 0, 'color' => 000000];
		// $footer->addLine($lineStyle);
		// $footer->addPreserveText('Page {PAGE} of {NUMPAGES}', null, $footerStyle);
		// $footer->addText(htmlspecialchars($company[4]), null, $footerStyle);

		// $headerStyle = ['size' => 11, 'bold' => true];
		// $paragraphStyle = ['align' => 'both', 'indent' => 0.5];

		// $productTab = 'productTab';
		// $phpWord->addParagraphStyle($productTab, ['tabs' => [new PhpWord\Style\Tab('left', 6000)] ]);
		// $section->addText(
		// 	htmlspecialchars('Product Name: '.$company[5].'	Product No: '.$company[6]), $headerStyle, $productTab
		// );
		// $section->addTextBreak(2);

		// $numberStyleList = $headerStyle + ['listType' => PhpWord\Style\ListItem::TYPE_NUMBER_NESTED];
		// $section->addListItem('SELECTION OF RISK MANAGEMENT STANDARD', 0, $headerStyle, $numberStyleList);
		// $section->addTextBreak(1);
		// $section->addText('The following standard is applicable to the Risk Management Plan of Axil Scientific Pte. Ltd.:', null, $paragraphStyle);
		// $section->addTextBreak(1);
		// $section->addText(htmlspecialchars($company[7]), null, $paragraphStyle);
		// $section->addTextBreak(1);
		// $section->addListItem('PURPOSE', 0, $headerStyle, $numberStyleList);
		// $section->addTextBreak(1);
		// $section->addText(htmlspecialchars($company[8]), null, $paragraphStyle);
		// $section->addTextBreak(1);
		// $section->addListItem('RISK MANAGEMENT ACTIVITIES', 0, $headerStyle, $numberStyleList);
		// $section->addTextBreak(1);
		// $section->addText(htmlspecialchars($company[9]), null, $paragraphStyle);
		// $section->addTextBreak(1);
		
		// $section->addText('SIGNATORY APPROVAL', $headerStyle);
		// $section->addTextBreak(1);
		// $tableName = 'Fancy Table';
		// $tableStyle = ['borderSize' => 6, 'borderColor' => '000000', 'cellMargin' => 80];
		// $cellVCenter = ['valign' => 'center'];
		// $cellHCenter = ['alignment' => PhpWord\SimpleType\JcTable::CENTER];
		// $cellBold = ['bold' => true];
		// $phpWord->addTableStyle($tableName, $tableStyle);
		// $table = $section->addTable($tableName);
		// $table->addRow(350);
		// $table->addCell(1820, $cellVCenter);
		// $table->addCell(1820, $cellVCenter)->addText("Name", $cellBold, $cellHCenter);
		// $table->addCell(1820, $cellVCenter)->addText("Designation", $cellBold, $cellHCenter);
		// $table->addCell(1820, $cellVCenter)->addText("Signature", $cellBold, $cellHCenter);
		// $table->addCell(1820, $cellVCenter)->addText("Date", $cellBold, $cellHCenter);
		// $table->addRow(900);
		// $table->addCell(1820)->addText("Prepared by:");
		// $table->addCell(1820)->addText(htmlspecialchars($company[10]), null, $cellHCenter);
		// $table->addCell(1820)->addText(htmlspecialchars($company[11]), null, $cellHCenter);
		// $table->addCell(1820);
		// $table->addCell(1820);
		// $table->addRow(900);
		// $table->addCell(1820)->addText("Approved by:");
		// $table->addCell(1820)->addText(htmlspecialchars($company[12]), null, $cellHCenter);
		// $table->addCell(1820)->addText(htmlspecialchars($company[13]), null, $cellHCenter);
		// $table->addCell(1820);
		// $table->addCell(1820);

		// $file = 'risk_mgmt.docx';
		// header("Content-Description: File Transfer");
		// header('Content-Disposition: attachment; filename="' . $file . '"');
		// header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
		// header('Content-Transfer-Encoding: binary');
		// header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
		// header('Expires: 0');
		// $xmlWriter = PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
		// return $xmlWriter->save("php://output");


		// $template = new NewTemplateProcessor('assets/template.docx');
		// $template->setImg('logo', ['src' => $request->file('logo'), 'swh'=>'100']);
		// foreach ($excel as $key => $data) {
		// 	$template->setValue(str_slug($data->question, '_'), htmlspecialchars($data->answer));
		// }


		$replace1 = "word/media/image1.png";
		$img1 = file_get_contents($request->file('logo'));
		$template = new OpenTemplateProcessor('assets/template2.docx');
		foreach ($excel as $key => $data) {
			$template->setValue(str_slug($data->question, '_'), htmlspecialchars($data->answer));
		}
		$template->zipClass->AddFromString($replace1, $img1);
		try {
			$template->saveAs(storage_path('risk_mgmt.docx'));
		} catch (Exception $e) {
		}
		return response()->download(storage_path('risk_mgmt.docx'))->deleteFileAfterSend(true);
	}

	public function csvData(Request $request)
	{
		$excel = Excel::load('assets/Sample Data.csv')->get();
		$result = [];
		foreach ($excel as $key => $data) {
			$result[str_slug($data->question, '_')] = $data->answer;
		}
		$result['logo'] = 'assets/images/Sample Logo.png';
		return response()->json($result);
	}
}
