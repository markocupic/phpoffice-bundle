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
$objPhpWord->replaceWithImage('my-best-image', 'files/best_image.jpg', array('width'=>'60mm'));


// Clone rows
// Push first datarecord to cloned row
$row = array(
    array('rank', '1', array('multiline' => false)),
    array('number', '501', array('multiline' => false)),
    array('firstname', 'James', array('multiline' => false)),
    array('lastname', 'First', array('multiline' => false)),
    array('time', '01:23:55', array('multiline' => false)),
    // Add an image with a predefined height
    array('my_image', 'files/image1.jpg', array('type' => 'image', 'height' => '50mm')),
);
$objPhpWord->replaceAndClone('rank', $row);

// Push second datarecord to cloned row
$row = array(
    array('rank', '1', array('multiline' => false)),
    array('number', '504', array('multiline' => false)),
    array('firstname', 'Niki', array('multiline' => false)),
    array('lastname', 'Last', array('multiline' => false)),
    array('time', '01:26:55', array('multiline' => false)),
    // Add an image with a predefined width
    array('my_image', 'files/image2.jpg', array('type' => 'image', 'width' => '50mm')),
);
$objPhpWord->replaceAndClone('rank', $row);

// Push third datarecord, etc...
$row = array(/** **/); 

// Create & send file to browser
$objPhpWord->sendToBrowser(true)
    ->generateUncached(true)
    ->generate();

```


