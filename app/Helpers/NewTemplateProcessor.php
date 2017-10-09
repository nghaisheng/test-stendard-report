<?php

namespace App\Helpers;

use PhpOffice\PhpWord\Settings;
use PhpOffice\PhpWord\Escaper\RegExp;
use PhpOffice\PhpWord\Escaper\Xml;
use PhpOffice\PhpWord\Exception\CopyFileException;
use PhpOffice\PhpWord\Exception\CreateTemporaryFileException;
use PhpOffice\PhpWord\Exception\Exception;
use PhpOffice\PhpWord\Shared\ZipArchive;
use Zend\Stdlib\StringUtils;
use PhpOffice\PhpWord\TemplateProcessor;

class NewTemplateProcessor extends TemplateProcessor
{
	protected $_rels;
	protected $_types;

	public function __construct($documentTemplate)
    {
        // Temporary document filename initialization
        $this->tempDocumentFilename = tempnam(Settings::getTempDir(), 'PhpWord');
        if (false === $this->tempDocumentFilename) {
            throw new CreateTemporaryFileException();
        }

        // Template file cloning
        if (false === copy($documentTemplate, $this->tempDocumentFilename)) {
            throw new CopyFileException($documentTemplate, $this->tempDocumentFilename);
        }

        // Temporary document content extraction
        $this->zipClass = new ZipArchive();
        $this->zipClass->open($this->tempDocumentFilename);
        $index = 1;
        while (false !== $this->zipClass->locateName($this->getHeaderName($index))) {
            $this->tempDocumentHeaders[$index] = $this->fixBrokenMacros(
                $this->zipClass->getFromName($this->getHeaderName($index))
            );
            $index++;
        }
        $index = 1;
        while (false !== $this->zipClass->locateName($this->getFooterName($index))) {
            $this->tempDocumentFooters[$index] = $this->fixBrokenMacros(
                $this->zipClass->getFromName($this->getFooterName($index))
            );
            $index++;
        }
        $this->tempDocumentMainPart = $this->fixBrokenMacros($this->zipClass->getFromName($this->getMainPartName()));
        $this->_countRels = 100;
    }

    public function save()
    {
        foreach ($this->tempDocumentHeaders as $index => $xml) {
            $this->zipClass->addFromString($this->getHeaderName($index), $xml);
        }

        $this->zipClass->addFromString($this->getMainPartName(), $this->tempDocumentMainPart);

        if ($this->_rels != "") {
			$this->zipClass->addFromString('word/_rels/document.xml.rels', $this->_rels);
		}
		if ($this->_types != "") {
			$this->zipClass->addFromString('[Content_Types].xml', $this->_types);
		}

        foreach ($this->tempDocumentFooters as $index => $xml) {
            $this->zipClass->addFromString($this->getFooterName($index), $xml);
        }

        // Close zip file
        if (false === $this->zipClass->close()) {
            throw new Exception('Could not close zip file.');
        }

        return $this->tempDocumentFilename;
    }

    function limpiarString($str) {
		return str_replace(
			array('&', '<', '>', "\n"), 
			array('&amp;', '&lt;', '&gt;', "\n" . '<w:br/>'), 
			$str
		);
	}

	public function setImg($strKey, $img) 
	{
		$strKey = '${'.$strKey.'}';
		$relationTmpl = '<Relationship Id="RID" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/IMG"/>';

		$imgTmpl = '<w:pict><v:shape type="#_x0000_t75" style="width:WIDpx;height:HEIpx"><v:imagedata r:id="RID" o:title=""/></v:shape></w:pict>';

		$toAdd = $toAddImg = $toAddType = '';
		$aSearch = array('RID', 'IMG');
		$aSearchType = array('IMG', 'EXT');
		$countrels = $this->_countRels++;
		//I'm work for jpg files, if you are working with other images types -> Write conditions here
		$imgExt = 'jpg';
		$imgName = 'img' . $countrels . '.' . $imgExt;

		$this->zipClass->deleteName('word/media/' . $imgName);
		$this->zipClass->addFile($img['src'], 'word/media/' . $imgName);

		$typeTmpl = '<Override PartName="/word/media/'.$imgName.'" ContentType="image/EXT"/>';

		$rid = 'rId' . $countrels;
		$countrels++;
		list($w,$h) = getimagesize($img['src']);

		if (isset($img['swh'])) { //Image proportionally larger side
			if($w <= $h) {
				$ht = (int)$img['swh'];
				$ot = $w/$h;
				$wh = (int)$img['swh']*$ot;
				$wh = round($wh);
			}
			if($w >= $h) {
				$wh = (int)$img['swh'];
				$ot = $h/$w;
				$ht = (int)$img['swh']*$ot;
				$ht = round($ht);
			}
			$w = $wh;
			$h = $ht;
		}

		if (isset($img['size'])) {
			$w = $img['size'][0];
			$h = $img['size'][1];           
		}

		$toAddImg .= str_replace(array('RID', 'WID', 'HEI'), array($rid, $w, $h), $imgTmpl);
		if (isset($img['dataImg'])) {
			$toAddImg.='<w:br/><w:t>'.$this->limpiarString($img['dataImg']).'</w:t><w:br/>';
		}

		$aReplace = array($imgName, $imgExt);
		$toAddType .= str_replace($aSearchType, $aReplace, $typeTmpl);

		$aReplace = array($rid, $imgName);
		$toAdd .= str_replace($aSearch, $aReplace, $relationTmpl);

		// $this->tempDocumentHeaders = str_replace('<w:t>' . $strKey . '</w:t>', $toAddImg, $this->tempDocumentHeaders);
		$this->tempDocumentMainPart = str_replace('<w:t>' . $strKey . '</w:t>', $toAddImg, $this->tempDocumentMainPart);
		// $this->tempDocumentFooters = str_replace('<w:t>' . $strKey . '</w:t>', $toAddImg, $this->tempDocumentFooters);

		if ($this->_rels == "") {
			$this->_rels = $this->zipClass->getFromName('word/_rels/document.xml.rels');
			$this->_types = $this->zipClass->getFromName('[Content_Types].xml');
		}
		$this->_types = str_replace('</Types>', $toAddType, $this->_types) . '</Types>';
		$this->_rels = str_replace('</Relationships>', $toAdd, $this->_rels) . '</Relationships>';
	}
}