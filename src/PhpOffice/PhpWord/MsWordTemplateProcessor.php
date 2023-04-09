<?php

declare(strict_types=1);

/*
 * This file is part of Php Office Bundle.
 *
 * (c) Marko Cupic 2023 <m.cupic@gmx.ch>
 * @license GPL-3.0-or-later
 * For the full copyright and license information,
 * please view the LICENSE file that was distributed with this source code.
 * @link https://github.com/markocupic/phpoffice-bundle
 */

namespace Markocupic\PhpOffice\PhpWord;

use Contao\System;
use PhpOffice\PhpWord\Exception\CopyFileException;
use PhpOffice\PhpWord\Exception\CreateTemporaryFileException;
use PhpOffice\PhpWord\TemplateProcessor;
use Symfony\Component\Filesystem\Exception\FileNotFoundException;
use Symfony\Component\Filesystem\Path;
use Symfony\Component\HttpFoundation\BinaryFileResponse;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\HttpFoundation\ResponseHeaderBag;
use Symfony\Component\Mime\MimeTypes;
use Symfony\Component\String\UnicodeString;

/**
 * @see README.md for usage explanations
 */
class MsWordTemplateProcessor extends TemplateProcessor
{
    public const ARR_DATA_CLONE_KEY = 'ARR_CLONES';
    public const ARR_DATA_REPLACEMENTS_KEY = 'ARR_REPLACEMENTS';
    protected string $templSrc;
    protected string $destinationSrc;
    protected array $arrData = [];
    protected bool $sendToBrowser = false;
    protected bool $sendToBrowserInline = false;
    protected bool $generateUncached = false;
    protected string|null $projectDir;

    /**
     * @throws CopyFileException
     * @throws CreateTemporaryFileException
     */
    public function __construct(string $templSrc, string $destinationSrc = '')
    {
        $this->projectDir = System::getContainer()->getParameter('kernel.project_dir');

        if ('' === $destinationSrc) {
            $destinationSrc = sprintf(
                sys_get_temp_dir().'/%s.docx',
                md5(microtime()).random_int(1000000, 9999999),
            );
        }

        $this->destinationSrc = Path::makeAbsolute($destinationSrc, $this->projectDir);
        $this->templSrc = Path::makeAbsolute($templSrc, $this->projectDir);

        if (!file_exists($this->templSrc)) {
            throw new FileNotFoundException(sprintf('Template file "%s" not found.', $this->templSrc));
        }

        $this->arrData = [static::ARR_DATA_REPLACEMENTS_KEY => [], static::ARR_DATA_CLONE_KEY => []];

        return parent::__construct($this->templSrc);
    }

    public function replace(string $search, $replace = '', array $options = []): void
    {
        $this->arrData[static::ARR_DATA_REPLACEMENTS_KEY][$search] = [
            'search' => (string) $search,
            'replace' => (string) $replace,
            'options' => $options,
        ];
    }

    public function replaceWithImage(string $search, $path = '', array $options = []): void
    {
        $path = Path::makeAbsolute($path,$this->projectDir);
        if (!is_file($path)) {
            return;
        }

        $arrImage = [
            'path' => $path,
        ];

        if (isset($options['width']) && '' !== $options['width']) {
            $arrImage['width'] = $options['width'];
            $arrImage['height'] = '';
        } elseif (isset($options['height']) && '' !== $options['height']) {
            $arrImage['height'] = $options['height'];
            $arrImage['width'] = '';
        }

        $limit = static::MAXIMUM_REPLACEMENTS_DEFAULT;

        if (isset($options['limit'])) {
            $limit = $options['limit'];
        }

        $this->setImageValue($search, $arrImage, $limit);
    }

    /**
     * Generate a new clone.
     */
    public function createClone(string $cloneKey): void
    {
        // Create new clone and push new row
        $this->arrData[static::ARR_DATA_CLONE_KEY][$cloneKey][] = [];
    }

    /**
     * @param string|int $search
     * @param string|int $replace
     */
    public function addToClone(string $cloneKey, $search, $replace = '', array $options = []): void
    {
        if (\is_array($this->arrData[static::ARR_DATA_CLONE_KEY][$cloneKey])) {
            $i = \count($this->arrData[static::ARR_DATA_CLONE_KEY][$cloneKey]) - 1;
            $this->arrData[static::ARR_DATA_CLONE_KEY][$cloneKey][$i][$search] = ['search' => $search, 'replace' => $replace, 'options' => $options];
        }
    }

    /**
     * @return $this
     */
    public function sendToBrowser(bool $blnSendToBrowser = false, bool $blnInline = false): self
    {
        $this->sendToBrowser = $blnSendToBrowser;
        $this->sendToBrowserInline = $blnInline;

        return $this;
    }

    /**
     * @return $this
     */
    public function generateUncached(bool $blnUncached = false): self
    {
        $this->generateUncached = $blnUncached;

        return $this;
    }

    /**
     * Generate the file.
     */
    public function generate(): Response|null
    {
        // Create docx file if it can not be found in the cache or if $this->generateUncached is set to true
        if (!is_file($this->destinationSrc) || true === $this->generateUncached) {
            // Process $this->arrData[static::ARR_DATA_CLONE_KEY] and replace the template vars
            foreach ($this->arrData[static::ARR_DATA_CLONE_KEY] as $cloneKey => $arrClones) {
                $countClones = \count($arrClones);

                if ($countClones > 0) {
                    // Clone rows
                    $this->cloneRow($cloneKey, $countClones);

                    $cloneIndex = 0;

                    foreach ($arrClones as $arrData) {
                        ++$cloneIndex;

                        foreach ($arrData as $replace) {
                            $replace['replace'] = htmlspecialchars(html_entity_decode((string) $replace['replace']));

                            // Format bold text (replace <B> & </B>)
                            $replace['replace'] = $this->formatBoldText((string) $replace['replace']);

                            // If multiline
                            if (isset($replace['options']['multiline']) && !empty($replace['options']['multiline'])) {
                                if (true === $replace['options']['multiline']) {
                                    $replace['replace'] = $this->formatMultilineText((string) $replace['replace']);
                                }
                            }

                            // If maximum replacement limit
                            if (!isset($replace['options']['limit'])) {
                                $replace['options']['limit'] = static::MAXIMUM_REPLACEMENTS_DEFAULT;
                            }

                            // Add image
                            if (isset($replace['replace']['type']) && 'image' === $replace['options']['type']) {
                                if (is_file($this->projectDir.'/'.$replace['replace'])) {
                                    $arrImg = [
                                        'path' => $this->projectDir.'/'.$replace['replace'],
                                        'height' => '',
                                        'width' => '',
                                    ];

                                    if (isset($replace['options']['width']) && '' !== $replace['options']['width']) {
                                        $arrImg['width'] = $replace['options']['width'];
                                    } elseif (isset($replace['options']['height']) && '' !== $replace['options']['height']) {
                                        $arrImg['height'] = $replace['options']['height'];
                                    }
                                    $this->setImageValue($replace['search'].'#'.$cloneIndex, $arrImg, $replace['options']['limit']);
                                }
                            } else { // Add text
                                $this->setValue($replace['search'].'#'.$cloneIndex, $replace['replace'], $replace['options']['limit']);
                            }
                        }
                    }
                }
            }

            // Process $this->arrData[static::ARR_DATA_REPLACEMENTS_KEY] and replace the template vars
            foreach ($this->arrData[static::ARR_DATA_REPLACEMENTS_KEY] as $replace) {
                $replace['replace'] = htmlspecialchars(html_entity_decode((string) $replace['replace']));

                // Format bold text (replace <B> & </B>)
                $replace['replace'] = $this->formatBoldText((string) $replace['replace']);

                // If multiline
                if (isset($replace['options']['multiline']) && !empty($replace['options']['multiline'])) {
                    if (true === $replace['options']['multiline']) {
                        $replace['replace'] = $this->formatMultilineText((string) $replace['replace']);
                    }
                }

                // If maximum replacement limit
                if (!isset($replace['options']['limit'])) {
                    $replace['options']['limit'] = static::MAXIMUM_REPLACEMENTS_DEFAULT;
                }

                $this->setValue($replace['search'], $replace['replace'], $replace['options']['limit']);
            }

            $this->saveAs($this->destinationSrc);
        }

        if ($this->sendToBrowser) {
            $fileName = basename($this->destinationSrc);

            return $this->binaryFileDownload($this->destinationSrc, $fileName, $this->sendToBrowserInline);
        }

        return null;
    }

    /**
     * Replace a block.
     * Overwrite original method.
     *
     * @param $blockname
     * @param $replacement
     */
    public function replaceBlock($blockname, $replacement): void
    {
        // Original pattern
        // '/(<\?xml.*)(<w:p.*>\${' . $blockname . '}<\/w:.*?p>)(.*)(<w:p.*\${\/' . $blockname . '}<\/w:.*?p>)/is',
        // Optimized pattern for Word 2017
        preg_match(
            '/(<\?xml.*)(<w:t.*>\${'.$blockname.'}<\/w:.*?t>)(.*)(<w:t.*\${\/'.$blockname.'}<\/w:.*?t>)/is',
            $this->tempDocumentMainPart,
            $matches
        );

        if (isset($matches[3])) {
            $this->tempDocumentMainPart = str_replace(
                $matches[2].$matches[3].$matches[4],
                $replacement,
                $this->tempDocumentMainPart
            );
        }
    }

    /**
     * Convert chars between <B> and </B> to bold text.
     */
    protected function formatBoldText(string $text): string
    {
        return str_replace(
            [
                '&lt;B&gt;',
                '&lt;/B&gt;',
            ],
            [
                '</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">',
                '</w:t></w:r><w:r><w:t xml:space="preserve">',
            ],
            $text
        );
    }

    /**
     * @param $text
     *
     * @return mixed|string
     */
    protected function formatMultilineText(string $text): string
    {
        return preg_replace('~\R~u', '</w:t><w:br/><w:t>', $text);
    }

    protected function binaryFileDownload(string $filePath, string $filename = '', bool $inline = false): Response
    {
        $response = new BinaryFileResponse($filePath);
        $response->setPrivate(); // public by default
        $response->setAutoEtag();

        $response->setContentDisposition(
            $inline ? ResponseHeaderBag::DISPOSITION_INLINE : ResponseHeaderBag::DISPOSITION_ATTACHMENT,
            $filename,
            (new UnicodeString(basename($filePath)))->ascii()->toString()
        );

        $mimeTypes = new MimeTypes();
        $mimeType = $mimeTypes->guessMimeType($filePath);

        $response->headers->addCacheControlDirective('must-revalidate');
        $response->headers->set('Connection', 'close');
        $response->headers->set('Content-Type', $mimeType);

        return $response->send();
    }
}
