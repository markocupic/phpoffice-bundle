# PHPOffice Bundle

## Generate easily Microsoft Word documents:

Watch the [demo template](https://github.com/markocupic/docx-from-template-bundle/blob/master/src/example/my_ms_word_template.docx)

```php
<?php
// Create phpword instance
$objPhpWord = Markocupic\PhpOffice\PhpWord\MsWordTemplateProcessor::create('vendor/markocupic/docx-from-template-bundle/src/example/my_ms_word_template.docx', 'system/tmp/output.docx');

// Options defaults
$optionsDefaults = array(
    'multiline' => false,
    'limit' => -1
);

// Simple replacement
$objPhpWord->pushData('category', 'Elite men');

// Another multiline replacement
$options = array('multiline' => true); 
$objPhpWord->pushData('sometext', 'Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt', $options);


// Clone rows
// Push first datarecord to cloned row
$row = array(
        array('key' => 'rank', 'value' => '1', 'options' => array('multiline' => false)),
        array('key' => 'number', 'value' => '501', 'options' => array('multiline' => false)),
        array('key' => 'firstname', 'value' => 'James', 'options' => array('multiline' => false)),
        array('key' => 'lastname', 'value' => 'Last', 'options' => array('multiline' => false)),
        array('key' => 'time', 'value' => '01:23:55', 'options' => array('multiline' => false)),
    );
$objPhpWord->pushClone('rank', $row);

// Push second datarecord to cloned row
$row = array(
    array('key' => 'rank', 'value' => '2', 'options' => array('multiline' => false)),
    array('key' => 'number', 'value' => '506', 'options' => array('multiline' => false)),
    array('key' => 'firstname', 'value' => 'Niki', 'options' => array('multiline' => false)),
    array('key' => 'lastname', 'value' => 'Nonsense', 'options' => array('multiline' => false)),
    array('key' => 'time', 'value' => '01:23:57', 'options' => array('multiline' => false)),
);
$objPhpWord->pushClone('rank', $row);

// Push third datarecord, etc...
$row = array(/** **/); 

// Create & send file to browser
$objPhpWord->sendToBrowser(true)
    ->generateUncached(true)
    ->generate();

```


