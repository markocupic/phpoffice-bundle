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
 * @see README.md for usage explanations
 * Class MsWordTemplateProcessor
 * @package Markocupic\PhpOffice\PhpWord
 */
class MsWordTemplateProcessor extends TemplateProcessor
{

    /**
     * @var array
     */
    protected $arrData = array();

    /**
     * Key name in sstatic::addData
     */
    const ARR_DATA_CLONE_KEY = 'ARR_CLONES';

    /**
     * Key name in sstatic::addData
     */
    const ARR_DATA_REPLACEMENTS_KEY = 'ARR_REPLACEMENTS';

    /**
     * @var
     */
    protected $templSrc;

    /**
     * @var
     */
    protected $destinationSrc;

    /**
     * @var bool
     */
    protected $sendToBrowser = false;

    /**
     * @var bool
     */
    protected $generateUncached = false;

    /**
     * @var
     */
    protected $rootDir;

    /**
     * MsWordTemplateProcessor constructor.
     * @param string $templSrc
     * @param string $destinationSrc
     * @throws \PhpOffice\PhpWord\Exception\CopyFileException
     * @throws \PhpOffice\PhpWord\Exception\CreateTemporaryFileException
     */
    public function __construct(string $templSrc, string $destinationSrc = '')
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

        $this->rootDir = $rootDir;
        $this->templSrc = $templSrc;
        $this->destinationSrc = $destinationSrc;
        $this->arrData = array(static::ARR_DATA_REPLACEMENTS_KEY => array(), static::ARR_DATA_CLONE_KEY => array());

        return parent::__construct($rootDir . '/' . $templSrc);
    }

    /**
     * @param string $search
     * @param string $replace
     * @param array $options
     */
    public function replace(string $search, $replace = '', array $options = array()): void
    {
        $this->arrData[static::ARR_DATA_REPLACEMENTS_KEY][$search] = array(
            'search'  => $search,
            'replace' => (string) $replace,
            'options' => $options
        );
    }

    /**
     * @param string $search
     * @param string $path
     * @param array $arrOptions
     */
    public function replaceWithImage(string $search, string $path = '', array $arrOptions)
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
    public function createClone(string $cloneKey): void
    {
        // Create new clone and push new row
        $this->arrData[static::ARR_DATA_CLONE_KEY][$cloneKey][] = array();
    }

    /**
     * To add data to a clone, you have to call first $this->createClone('cloneKey')
     * @param string $cloneKey
     * @param $search
     * @param $replace
     * @param $options
     */
    public function addToClone(string $cloneKey, $search, $replace = '', $options): void
    {
        if (is_array($this->arrData[static::ARR_DATA_CLONE_KEY][$cloneKey]))
        {
            $i = count($this->arrData[static::ARR_DATA_CLONE_KEY][$cloneKey]) - 1;
            $this->arrData[static::ARR_DATA_CLONE_KEY][$cloneKey][$i][$search] = array('search' => $search, 'replace' => $replace, 'options' => $options);
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
            // Process $this->arrData[static::ARR_DATA_CLONE_KEY] and replace the template vars
            foreach ($this->arrData[static::ARR_DATA_CLONE_KEY] as $cloneKey => $arrClones)
            {
                $countClones = count($arrClones);
                if ($countClones > 0)
                {
                    // Clone rows
                    $this->cloneRow($cloneKey, $countClones);

                    $cloneIndex = 0;
                    foreach ($arrClones as $arrData)
                    {
                        $cloneIndex++;

                        foreach ($arrData as $search => $replace)
                        {
                            // If multiline
                            if (isset($replace['options']['multiline']) && !empty($replace['options']['multiline']))
                            {
                                if ($replace['options']['multiline'] === true)
                                {
                                    $replace['replace'] = $this->formatMultilineText($replace['replace']);
                                }
                            }

                            // If maximum replacement limit
                            if (!isset($replace['options']['limit']))
                            {
                                $replace['options']['limit'] = static::MAXIMUM_REPLACEMENTS_DEFAULT;
                            }

                            // Add image
                            if (isset($replace['replace']['type']) && $replace['options']['type'] === 'image')
                            {
                                if (is_file($this->rootDir . '/' . $replace['replace']))
                                {
                                    $arrImg = array(
                                        'path'   => $this->rootDir . '/' . $replace['replace'],
                                        'height' => '',
                                        'width'  => ''
                                    );
                                    if (isset($replace['options']['width']) && $replace['options']['width'] != '')
                                    {
                                        $arrImg['width'] = $replace['options']['width'];
                                    }
                                    elseif (isset($replace['options']['height']) && $replace['options']['height'] != '')
                                    {
                                        $arrImg['height'] = $replace['options']['height'];
                                    }
                                    $this->setImageValue($replace['search'] . '#' . $cloneIndex, $arrImg, $replace['options']['limit']);
                                }
                            }
                            else // Add text
                            {
                                $this->setValue($replace['search'] . '#' . $cloneIndex, $replace['replace'], $replace['options']['limit']);
                            }
                        }
                    }
                }
            }

            // Process $this->arrData[static::ARR_DATA_REPLACEMENTS_KEY] and replace the template vars
            foreach ($this->arrData[static::ARR_DATA_REPLACEMENTS_KEY] as $search => $replace)
            {
                // If multiline
                if (isset($replace['options']['multiline']) && !empty($replace['options']['multiline']))
                {
                    if ($replace['options']['multiline'] === true)
                    {
                        $replace['replace'] = $this->formatMultilineText($replace['replace']);
                    }
                }

                // If maximum replacement limit
                if (!isset($replace['options']['limit']))
                {
                    $replace['options']['limit'] = static::MAXIMUM_REPLACEMENTS_DEFAULT;
                }

                $this->setValue($replace['search'], $replace['replace'], $replace['options']['limit']);
            }

            $this->saveAs($this->rootDir . '/' . $this->destinationSrc);
        }

        if ($this->sendToBrowser)
        {
            $objFile = new File($this->destinationSrc);
            $objFile->sendToBrowser();
        }
    }

    /**
     * @param $text
     * @return mixed|string
     */
    protected function formatMultilineText($text): string
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
