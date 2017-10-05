<?php

namespace App\Http\Controllers;

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
		$company = [];
		foreach ($excel as $key => $data) {
			$company[$data->qns_id] = $data->answer;
		}
		
		$phpWord = new PhpWord\PhpWord();
		$section = $phpWord->addSection();
		$headerStyle = ['alignment' => PhpWord\SimpleType\Jc::RIGHT];
		$header = $section->addHeader();
		$header->addImage($request->file('logo'), [
			'positioning' => 'relative',
			'height' => 40,
			'wrappingStyle' => 'infront',
		]);
		$header->addText($company[2].' '.$company[3], null, $headerStyle);
		$lineStyle = ['weight' => 1.5, 'width' => 600, 'height' => 0, 'color' => 000000];
		$header->addLine($lineStyle);

		$footerStyle = ['alignment' => PhpWord\SimpleType\Jc::RIGHT];
		$footer = $section->addFooter();
		$lineStyle = ['weight' => 1, 'width' => 600, 'height' => 0, 'color' => 000000];
		$footer->addLine($lineStyle);
		$footer->addPreserveText('Page {PAGE} of {NUMPAGES}', null, $footerStyle);
		$footer->addText($company[4], null, $footerStyle);

		$headerStyle = ['size' => 11, 'bold' => true];
		$paragraphStyle = ['align' => 'both', 'indent' => 0.5];

		$productTab = 'productTab';
		$phpWord->addParagraphStyle($productTab, ['tabs' => [new PhpWord\Style\Tab('left', 6000)] ]);
		$section->addText(
			'Product Name: '.$company[5].'	Product No: '.$company[6], $headerStyle, $productTab
		);
		$section->addTextBreak(2);

		$numberStyleList = $headerStyle + ['listType' => PhpWord\Style\ListItem::TYPE_NUMBER_NESTED];
		$section->addListItem('SELECTION OF RISK MANAGEMENT STANDARD', 0, $headerStyle, $numberStyleList);
		$section->addTextBreak(1);
		$section->addText('The following standard is applicable to the Risk Management Plan of Axil Scientific Pte. Ltd.:', null, $paragraphStyle);
		$section->addTextBreak(1);
		$section->addText($company[7], null, $paragraphStyle);
		$section->addTextBreak(1);
		$section->addListItem('PURPOSE', 0, $headerStyle, $numberStyleList);
		$section->addTextBreak(1);
		$section->addText($company[8], null, $paragraphStyle);
		$section->addTextBreak(1);
		$section->addListItem('RISK MANAGEMENT ACTIVITIES', 0, $headerStyle, $numberStyleList);
		$section->addTextBreak(1);
		$section->addText($company[9], null, $paragraphStyle);
		$section->addTextBreak(1);
		
		$section->addText('SIGNATORY APPROVAL', $headerStyle);
		$section->addTextBreak(1);
		$tableName = 'Fancy Table';
		$tableStyle = ['borderSize' => 6, 'borderColor' => '000000', 'cellMargin' => 80];
		$cellVCenter = ['valign' => 'center'];
		$cellHCenter = ['alignment' => PhpWord\SimpleType\JcTable::CENTER];
		$cellBold = ['bold' => true];
		$phpWord->addTableStyle($tableName, $tableStyle);
		$table = $section->addTable($tableName);
		$table->addRow(350);
		$table->addCell(1820, $cellVCenter);
		$table->addCell(1820, $cellVCenter)->addText("Name", $cellBold, $cellHCenter);
		$table->addCell(1820, $cellVCenter)->addText("Designation", $cellBold, $cellHCenter);
		$table->addCell(1820, $cellVCenter)->addText("Signature", $cellBold, $cellHCenter);
		$table->addCell(1820, $cellVCenter)->addText("Date", $cellBold, $cellHCenter);
		$table->addRow(900);
		$table->addCell(1820)->addText("Prepared by:");
		$table->addCell(1820)->addText($company[10], null, $cellHCenter);
		$table->addCell(1820)->addText($company[11], null, $cellHCenter);
		$table->addCell(1820);
		$table->addCell(1820);
		$table->addRow(900);
		$table->addCell(1820)->addText("Approved by:");
		$table->addCell(1820)->addText($company[12], null, $cellHCenter);
		$table->addCell(1820)->addText($company[13], null, $cellHCenter);
		$table->addCell(1820);
		$table->addCell(1820);

		$file = 'risk_mgmt.docx';
		header("Content-Description: File Transfer");
		header('Content-Disposition: attachment; filename="' . $file . '"');
		header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
		header('Content-Transfer-Encoding: binary');
		header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
		header('Expires: 0');
		$xmlWriter = PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
		return $xmlWriter->save("php://output");

		// $template = new PhpWord\TemplateProcessor('assets/template.docx');
		// $template->setImg('logo', array('src' => $request->file('logo'),'swh'=>'100'));
		// $template->setValue('document_no', $company[2]);
		// $template->setValue('document_title', $company[3]);
		// $template->setValue('version_no', $company[4]);
		// $template->setValue('product_name', $company[5]);
		// $template->setValue('product_no', $company[6]);
		// $template->setValue('risk_mgmt_standard', $company[7]);
		// $template->setValue('purpose', $company[8]);
		// $template->setValue('risk_mgmt_activities', $company[9]);
		// $template->setValue('prepared_by', $company[10]);
		// $template->setValue('designation_preparer', $company[11]);
		// $template->setValue('approved_by', $company[12]);
		// $template->setValue('designation_approver', $company[13]);
		// try {
		// 	$template->saveAs(storage_path('risk_mgmt.docx'));
		// } catch (Exception $e) {
		// }
		// return response()->download(storage_path('risk_mgmt.docx'))->deleteFileAfterSend(true);
	}
}
