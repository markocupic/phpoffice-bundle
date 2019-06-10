<?php
/**
 * PhpOffice helper classes
 * Copyright (c) 2008-2019 Marko Cupic
 * @package phpoffice-bundle
 * @author Marko Cupic m.cupic@gmx.ch, 2019
 * @link https://github.com/markocupic/phpoffice-bundle
 */

declare(strict_types=1);

namespace Markocupic\PhpOffice\PhpWord;

use Contao\File;
use Contao\System;
use Symfony\Component\Filesystem\Exception\FileNotFoundException;
use PhpOffice\PhpWord\TemplateProcessor;

/**
 * @see README.md
 * Class MsWordTemplateProcessor
 * @package Markocupic\PhpOffice\PhpWord
 */
class MsWordTemplateProcessor extends TemplateProcessor
{

    /**
     * @var array
     */
    private $arrData = array();

    /**
     * @var
     */
    private $templSrc;

    /**
     * @var
     */
    private $destinationSrc;

    /**
     * @var bool
     */
    private $sendToBrowser = false;

    /**
     * @var bool
     */
    private $generateUncached = false;

    /**
     * @var
     */
    private $rootDir;

    /**
     * @param string $templSrc
     * @param string $destinationSrc
     * @return GenerateDocxFromTemplate
     * @throws \PhpOffice\PhpWord\Exception\CopyFileException
     * @throws \PhpOffice\PhpWord\Exception\CreateTemporaryFileException
     */
    public static function create(string $templSrc, string $destinationSrc = '')
    {
        if ($destinationSrc === '')
        {
            $destinationSrc = sprintf('system/tmp/%s.docx', md5(microtime()) . rand(1000000, 9999999));
        }

        $rootDir = System::getContainer()->getParameter('kernel.project_dir');
        if (!file_exists($rootDir . '/' . $templSrc))
        {
            throw new FileNotFoundException(sprintf('Template file "%s" not found.', $templSrc));
        }

        $self = new static($rootDir . '/' . $templSrc);
        $self->rootDir = $rootDir;
        $self->templSrc = $templSrc;
        $self->destinationSrc = $destinationSrc;
        return $self;
    }

    /**
     * @param string $key
     * @param $value
     * @param array $options
     */
    public function replace(string $key, $value, array $options = array())
    {
        if (!is_array($this->arrData))
        {
            $this->arrData = [];
        }

        foreach ($this->arrData as $k => $v)
        {
            if ($v['key'] === $key)
            {
                $this->arrData[$k]['value'] = $value;
                $this->arrData[$k]['options'] = $options;
                return;
            }
        }
        $this->arrData[] = array(
            'key'     => $key,
            'value'   => $value,
            'options' => $options
        );
    }

    /**
     * @param string $search
     * @param string $path
     * @param array $arrOptions
     */
    public function replaceWithImage(string $search, string $path, array $arrOptions)
    {
        if (!is_file($this->rootDir . '/' . $path))
        {
            return;
        }

        $arrImage = array(
            'path' => $this->rootDir . '/' . $path
        );

        if (isset($arrOptions['width']) && $arrOptions['width'] != '')
        {
            $arrImage['width'] = $arrOptions['width'];
            $arrImage['height'] = '';
        }
        elseif (isset($arrOptions['height']) && $arrOptions['height'] != '')
        {
            $arrImage['height'] = $arrOptions['height'];
            $arrImage['width'] = '';
        }

        $limit = static::MAXIMUM_REPLACEMENTS_DEFAULT;
        if (isset($arrOptions['limit']))
        {
            $limit = $arrOptions['limit'];
        }

        $this->setImageValue($search, $arrImage, $limit);
    }

    /**
     * Generate a new clone
     * @param string $cloneKey
     */
    public function createClone(string $cloneKey)
    {
        // Check if clone already exists
        foreach ($this->arrData as $k => $v)
        {
            if ($this->arrData[$k]['clone'] === $cloneKey)
            {
                // Push new row
                $this->arrData[$k]['rows'][] = array();
                return;
            }
        }
        // Create new clone and push new row
        $this->arrData[] = array(
            'clone' => $cloneKey,
            'rows'  => array(array())
        );
    }

    /**
     * To add data to a clone, you have to call first $this->createClone('cloneKey')
     * @param string $cloneKey
     * @param $search
     * @param $replace
     * @param $options
     */
    public function addToClone(string $cloneKey, $search, $replace, $options)
    {
        // Check if clone already exists
        foreach ($this->arrData as $k => $v)
        {
            if ($this->arrData[$k]['clone'] === $cloneKey)
            {
                if (is_array($this->arrData[$k]['rows']))
                {
                    $i = count($this->arrData[$k]['rows']) - 1;
                    $this->arrData[$k]['rows'][$i][] = array('key' => $search, 'value' => $replace, 'options' => $options);
                }
            }
        }
    }

    /**
     * @param bool $blnSendToBrowser
     * @return static
     */
    public function sendToBrowser($blnSendToBrowser = false): self
    {
        $this->sendToBrowser = $blnSendToBrowser;
        return $this;
    }

    /**
     * @param bool $blnUncached
     * @return static
     */
    public function generateUncached($blnUncached = false): self
    {
        $this->generateUncached = $blnUncached;
        return $this;
    }

    /**
     * Generate the file
     */
    public function generate(): void
    {
        // Create docx file if it can not be found in the cache or if $this->generateUncached is set to true
        if (!is_file($this->rootDir . '/' . $this->destinationSrc) || $this->generateUncached === true)
        {
            // Process $this->arrData and replace the template vars
            foreach ($this->arrData as $aData)
            {
                if (isset($aData['clone']) && !empty($aData['clone']))
                {
                    // Clone rows
                    if (count($aData['rows']) > 0)
                    {
                        $this->cloneRow($aData['clone'], count($aData['rows']));

                        $row = 0;
                        foreach ($aData['rows'] as $key => $arrRow)
                        {
                            $row = $key + 1;
                            foreach ($arrRow as $arrRowData)
                            {
                                // If multiline
                                if (isset($arrRowData['options']['multiline']) && !empty($arrRowData['options']['multiline']))
                                {
                                    if ($arrRowData['options']['multiline'] === true)
                                    {
                                        $arrRowData['value'] = static::formatMultilineText($arrRowData['value']);
                                    }
                                }

                                // If maximum replacement limit
                                if (!isset($arrRowData['options']['limit']))
                                {
                                    $arrRowData['options']['limit'] = static::MAXIMUM_REPLACEMENTS_DEFAULT;
                                }

                                // Add image
                                if (isset($arrRowData['value']) && $arrRowData['options']['type'] === 'image')
                                {
                                    if (is_file($this->rootDir . '/' . $arrRowData['value']))
                                    {
                                        $arrImg = array(
                                            'path'   => $this->rootDir . '/' . $arrRowData['value'],
                                            'height' => '',
                                            'width'  => ''
                                        );
                                        if (isset($arrRowData['options']['width']) && $arrRowData['options']['width'] != '')
                                        {
                                            $arrImg['width'] = $arrRowData['options']['width'];
                                        }
                                        elseif (isset($arrRowData['options']['height']) && $arrRowData['options']['height'] != '')
                                        {
                                            $arrImg['height'] = $arrRowData['options']['height'];
                                        }
                                        $this->setImageValue($arrRowData['key'] . '#' . $row, $arrImg, $arrRowData['options']['limit']);
                                    }
                                }
                                else // Add text
                                {
                                    $this->setValue($arrRowData['key'] . '#' . $row, $arrRowData['value'], $arrRowData['options']['limit']);
                                }
                            }
                        }
                    }
                }
                else
                {
                    // If multiline
                    if (isset($aData['options']['multiline']) && !empty($aData['options']['multiline']))
                    {
                        if ($aData['options']['multiline'] === true)
                        {
                            $aData['value'] = static::formatMultilineText($aData['value']);
                        }
                    }

                    // If maximum replacement limit
                    if (!isset($aData['options']['limit']))
                    {
                        $aData['options']['limit'] = static::MAXIMUM_REPLACEMENTS_DEFAULT;
                    }

                    $this->setValue($aData['key'], $aData['value'], $aData['options']['limit']);
                }
            }

            $this->saveAs($this->rootDir . '/' . $this->destinationSrc);
        }

        if ($this->sendToBrowser)
        {
            $objDocx = new File($this->destinationSrc);
            $objDocx->sendToBrowser();
        }
    }

    /**
     * @param $text
     * @return mixed|string
     */
    protected static function formatMultilineText($text): string
    {
        $text = htmlspecialchars(html_entity_decode($text));
        $text = preg_replace('~\R~u', '</w:t><w:br/><w:t>', $text);
        return $text;
    }

    /**
     * Replace a block.
     * Overwrite original method
     * @param string $blockname
     * @param string $replacement
     *
     * @return void
     */
    public function replaceBlock($blockname, $replacement)
    {
        // Original pattern
        // '/(<\?xml.*)(<w:p.*>\${' . $blockname . '}<\/w:.*?p>)(.*)(<w:p.*\${\/' . $blockname . '}<\/w:.*?p>)/is',
        // Optimized pattern for Word 2017
        preg_match(

            '/(<\?xml.*)(<w:t.*>\${' . $blockname . '}<\/w:.*?t>)(.*)(<w:t.*\${\/' . $blockname . '}<\/w:.*?t>)/is',
            $this->tempDocumentMainPart,
            $matches
        );

        if (isset($matches[3]))
        {
            $this->tempDocumentMainPart = str_replace(
                $matches[2] . $matches[3] . $matches[4],
                $replacement,
                $this->tempDocumentMainPart
            );
        }
    }

}
