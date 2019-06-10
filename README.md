# PHPOffice Bundle

## MsWordTemplateProcessor
#### Generate easily Microsoft Word documents:

Watch the [demo template](https://github.com/markocupic/phpoffice-bundle/blob/master/src/howto/msword_template_processor.docx)

```php
<?php
// Create phpword instance
$objPhpWord = Markocupic\PhpOffice\PhpWord\MsWordTemplateProcessor::create('vendor/markocupic/docx-from-template-bundle/src/example/my_ms_word_template.docx', 'system/tmp/output.docx');

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
$objPhpWord->replaceWithImage('my-best-image', 'files/best_image.jpg', array('width' => '60mm'));

// Clone rows
// Push first datarecord to cloned row
$objPhpWord->createClone('rank');
$objPhpWord->addToClone('rank', 'rank', '1', array('multiline' => false));
$objPhpWord->addToClone('rank', 'number', '501', array('multiline' => false));
$objPhpWord->addToClone('rank', 'firstname', 'James', array('multiline' => false));
$objPhpWord->addToClone('rank', 'lastname', 'First', array('multiline' => false));
$objPhpWord->addToClone('rank', 'time', '01:23:55', array('multiline' => false));
// Add an image with a predefined height
$objPhpWord->addToClone('rank', 'my_image', 'files/image1.jpg', array('type' => 'image', 'height' => '50mm'));

// Push second datarecord to cloned row
$objPhpWord->createClone('rank');
$objPhpWord->addToClone('rank', 'rank', '2', array('multiline' => false));
$objPhpWord->addToClone('rank', 'number', '503', array('multiline' => false));
$objPhpWord->addToClone('rank', 'firstname', 'James', array('multiline' => false));
$objPhpWord->addToClone('rank', 'lastname', 'Last', array('multiline' => false));
$objPhpWord->addToClone('rank', 'time', '01:25:55', array('multiline' => false));
// Add an image with a predefined width
$objPhpWord->addToClone('rank', 'my_image', 'files/image2.jpg', array('type' => 'image', 'width' => '70mm'));

// Push third datarecord, etc...
$objPhpWord->createClone('rank');
// .... etc.

// Create & send file to browser
$objPhpWord->sendToBrowser(true)
    ->generateUncached(true)
    ->generate();

```


