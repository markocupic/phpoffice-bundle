<?php

declare(strict_types=1);

/*
 * This file is part of Php Office Bundle.
 *
 * (c) Marko Cupic 2024 <m.cupic@gmx.ch>
 * @license MIT
 * For the full copyright and license information,
 * please view the LICENSE file that was distributed with this source code.
 * @link https://github.com/markocupic/phpoffice-bundle
 */

namespace Markocupic\PhpOfficeBundle\Tests\PhpOffice\PhpWord;

use Markocupic\PhpOffice\PhpWord\MsWordTemplateProcessor;
use PHPUnit\Framework\TestCase;
use Symfony\Component\Filesystem\Exception\FileNotFoundException;
use Symfony\Component\Filesystem\Filesystem;

class MsWordTemplateProcessorTest extends TestCase
{
    private const EXAMPLE_DIR = __DIR__.'/../../../docs/example';

    private string $target;

    protected function setUp(): void
    {
        parent::setUp();

        $this->target = sys_get_temp_dir().'/phpoffice-bundle-test-'.bin2hex(random_bytes(6)).'.docx';
    }

    protected function tearDown(): void
    {
        (new Filesystem())->remove($this->target);

        parent::tearDown();
    }

    public function testThrowsIfTheTemplateDoesNotExist(): void
    {
        $this->expectException(FileNotFoundException::class);

        new MsWordTemplateProcessor(self::EXAMPLE_DIR.'/templates/does_not_exist.docx', $this->target);
    }

    public function testGeneratesTheDocument(): void
    {
        $processor = new MsWordTemplateProcessor(self::EXAMPLE_DIR.'/templates/ms_word_template.docx', $this->target);

        $processor->replace('category', "<B>Elite</B> men & women\nline 2", ['multiline' => true]);
        $processor->replaceWithImage('my-best-image', self::EXAMPLE_DIR.'/assets/my-best-image.jpg', ['width' => '160mm']);

        foreach ([['1', 'James', 'First', 'avatar_1.png'], ['2', 'Jane', 'Last', 'avatar_2.png']] as [$rank, $firstname, $lastname, $avatar]) {
            $processor->createClone('rank');
            $processor->addToClone('rank', 'rank', $rank);
            $processor->addToClone('rank', 'number', '50'.$rank);
            $processor->addToClone('rank', 'firstname', $firstname);
            $processor->addToClone('rank', 'lastname', $lastname);
            $processor->addToClone('rank', 'time', '01:23:55');
            $processor->addToClone('rank', 'avatar', self::EXAMPLE_DIR.'/assets/'.$avatar, ['type' => 'image', 'height' => '30mm']);
        }

        $this->assertCount(2, $processor->getData()[MsWordTemplateProcessor::ARR_DATA_CLONE_KEY]['rank']);

        $file = $processor->generate();

        $this->assertSame($this->target, $file->getPathname());
        $this->assertFileExists($this->target);

        $zip = new \ZipArchive();
        $this->assertTrue($zip->open($this->target));

        $xml = (string) $zip->getFromName('word/document.xml');

        $media = 0;

        for ($i = 0; $i < $zip->numFiles; ++$i) {
            if (str_starts_with((string) $zip->getNameIndex($i), 'word/media/')) {
                ++$media;
            }
        }

        $zip->close();

        // Text replacements are XML escaped, bold and multiline formatting is applied
        $this->assertStringContainsString('<w:b/></w:rPr><w:t xml:space="preserve">Elite', $xml);
        $this->assertStringContainsString(' men &amp; women</w:t><w:br/><w:t>line 2', $xml);
        $this->assertStringNotContainsString('category', $xml);

        // Cloned rows
        $this->assertStringContainsString('James', $xml);
        $this->assertStringContainsString('Jane', $xml);
        $this->assertStringNotContainsString('firstname', $xml);

        // Images in cloned rows are replaced as well (existing image + main image + 2 avatars)
        $this->assertStringNotContainsString('avatar', $xml);
        $this->assertGreaterThanOrEqual(4, $media);
    }
}
