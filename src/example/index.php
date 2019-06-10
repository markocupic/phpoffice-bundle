<?php

// Create phpword instance
$objPhpWord = new Markocupic\PhpOffice\PhpWord\MsWordTemplateProcessor('vendor/markocupic/phpoffice-bundle/src/example/templates/ms_word_template.docx', 'system/tmp/result.docx');

// Options defaults
//$optionsDefaults = array(
//    'multiline' => false,
//    'limit' => -1
//);

// Simple replacement
$objPhpWord->replace('category', 'Elite men');

// Another multiline replacement
$options = array('multiline' => true);
$objPhpWord->replace('sometext', 'Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt', $options);

// Image replacement
$objPhpWord->replaceWithImage('my-best-image', 'vendor/markocupic/phpoffice-bundle/src/example/assets/my-best-image.jpg', array('width' => '160mm'));

// Clone rows
// Push first datarecord to cloned row
$objPhpWord->createClone('rank');
$objPhpWord->addToClone('rank', 'rank', '1', array('multiline' => false));
$objPhpWord->addToClone('rank', 'number', '501', array('multiline' => false));
$objPhpWord->addToClone('rank', 'firstname', 'James', array('multiline' => false));
$objPhpWord->addToClone('rank', 'lastname', 'First', array('multiline' => false));
$objPhpWord->addToClone('rank', 'time', '01:23:55', array('multiline' => false));
// Add an image with a predefined height
$objPhpWord->addToClone('rank', 'avatar', 'vendor/markocupic/phpoffice-bundle/src/example/assets/avatar_1.png', array('type' => 'image', 'height' => '30mm'));

// Push second datarecord to cloned row
$objPhpWord->createClone('rank');
$objPhpWord->addToClone('rank', 'rank', '2', array('multiline' => false));
$objPhpWord->addToClone('rank', 'number', '503', array('multiline' => false));
$objPhpWord->addToClone('rank', 'firstname', 'James', array('multiline' => false));
$objPhpWord->addToClone('rank', 'lastname', 'Last', array('multiline' => false));
$objPhpWord->addToClone('rank', 'time', '01:25:55', array('multiline' => false));
// Add an image with a predefined width
$objPhpWord->addToClone('rank', 'avatar', 'vendor/markocupic/phpoffice-bundle/src/example/assets/avatar_2.png', array('type' => 'image', 'width' => '28.3mm'));

// Push third datarecord, etc...
//$objPhpWord->createClone('rank');
// .... etc.

// Create & send file to browser
$objPhpWord->sendToBrowser(true)
    ->generateUncached(true)
    ->generate();
