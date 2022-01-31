<?php

declare(strict_types=1);

/*
 * This file is part of Php Office Bundle.
 *
 * (c) Marko Cupic 2022 <m.cupic@gmx.ch>
 * @license GPL-3.0-or-later
 * For the full copyright and license information,
 * please view the LICENSE file that was distributed with this source code.
 * @link https://github.com/markocupic/phpoffice-bundle
 */

use Markocupic\PhpOffice\PhpWord\MsWordTemplateProcessor;

$objPhpWord = new MsWordTemplateProcessor('vendor/markocupic/phpoffice-bundle/src/example/templates/ms_word_template.docx', 'system/tmp/result.docx');

// Simple replacement
$objPhpWord->replace('category', 'Elite men');

// Another multiline replacement
$options = ['multiline' => true];
$objPhpWord->replace('sometext', 'Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt', $options);

// Image replacement
$objPhpWord->replaceWithImage('my-best-image', 'vendor/markocupic/phpoffice-bundle/src/example/assets/my-best-image.jpg', ['width' => '160mm']);

// Clone rows
// Push first datarecord to cloned row
$objPhpWord->createClone('rank');
$objPhpWord->addToClone('rank', 'rank', '1', ['multiline' => false]);
$objPhpWord->addToClone('rank', 'number', '501', ['multiline' => false]);
$objPhpWord->addToClone('rank', 'firstname', 'James', ['multiline' => false]);
$objPhpWord->addToClone('rank', 'lastname', 'First', ['multiline' => false]);
$objPhpWord->addToClone('rank', 'time', '01:23:55', ['multiline' => false]);

// Add an image with a predefined height
$objPhpWord->addToClone('rank', 'avatar', 'vendor/markocupic/phpoffice-bundle/src/example/assets/avatar_1.png', ['type' => 'image', 'height' => '30mm']);

// Push second datarecord to cloned row
$objPhpWord->createClone('rank');
$objPhpWord->addToClone('rank', 'rank', '2', ['multiline' => false]);
$objPhpWord->addToClone('rank', 'number', '503', ['multiline' => false]);
$objPhpWord->addToClone('rank', 'firstname', 'James', ['multiline' => false]);
$objPhpWord->addToClone('rank', 'lastname', 'Last', ['multiline' => false]);
$objPhpWord->addToClone('rank', 'time', '01:25:55', ['multiline' => false]);

// Add an image with a predefined width
$objPhpWord->addToClone('rank', 'avatar', 'vendor/markocupic/phpoffice-bundle/src/example/assets/avatar_2.png', ['type' => 'image', 'width' => '28.3mm']);

// Push third datarecord, etc...
//$objPhpWord->createClone('rank');
// .... etc.

// Create & send file to browser
$objPhpWord->sendToBrowser(true)
    ->generateUncached(true)
    ->generate()
;
