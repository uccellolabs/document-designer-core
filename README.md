Document Designer is a [Laravel](https://laravel.com) package allowing to generate documents from a template and filling it with some user data. The current version of PHPWord supports Microsoft Office Open XML (OOXML or OpenXML)

Document Designer is based on [PHPWord](https://github.com/PHPOffice/PHPWord).
[LibreOffce](https://www.libreoffice.org) need to be installed on the server for PDF export.

## Features

- Template processing from .docx files
- Template processing from .xslx files
- Export in .docx, .xlsx & .pdf fomat
- Text variable replacement
- Table row repetition
- Recursive blocks repetition



## Installation

#### Package
```bash
composer require uccello/document-designer-core
```


#### LibreOffice
You can refer to the [oficial documentation from LibreOffice](https://www.libreoffice.org/get-help/install-howto/) for the installation on your server OS.



## Getting Started

You just need to specify a template file, a out file name and the dataset to be used to parse and populate the template:

```php
DocumentIO::process($templateFile, $outFile, $data);
```

The `$data`  should be an associative array containing all the variables and the corresponding data to be replaced.

Depending the extension (.docx or .pdf) given in the `$outFile`, the export format will be DOCX or PDF.



#### Variables

```php
$data = [
    'variableName' => 'Content of the variable',
    'otherVariable' => 'Other content',
]
```

And in the template document, you need to declare the variables with the syntax : **${variableName}**



#### Tables

Table keys in the associative array needs a `t:` prefix.

```php
$data = [
    't:variableName' => [
        [
            'variableName' => 'dolor',
            'otherVariable' => 'elit',
        ],
        [
            'variableName' => 'amet',
            'otherVariable' => 'elit',
        ],
    ],
]
```

In the template document, you need to declare the variables with the syntax : **${variableName}**

The first row of the all template document containing a variable with the the same name as the table key will be repeated and the content of the other variables will be replaced.



#### Images

Image keys in the associative array needs a `i:` prefix.

```php
$data = [
    'i:imgVariable' => 'path/image.jpg',
]
```

And in the template document, you need to declare the variables with the syntax : **${imgVariable:[width]:[height]:[ratio]}**



#### Blocks

Block keys in the associative array needs a `b:` prefix and with capital letters.

```php
$data = [
    'b:BLOCK_NAME' => [
        [
            'variableName' => 'dolor',
            'otherVariable' => 'elit',
        ],
        [
            'variableName' => 'amet',
            'otherVariable' => 'elit',
        ],
    ]
]
```

In the template document, you need to declare the blocks start and end with flags :

**${BLOCK_NAME}**

Block content...

**${/BLOCK_NAME}**



The blocks behave recursively, witch means they can contain variables, tables and others block.

If a block contain a variable with a name in conflict with another variable in a parent block, the deeper block variables will be replaced in priority.



### Exemple:

##### Php

```php
use Uccello\DocumentDesignerCore\Support\DocumentIO;

$templateFile = "path/template.docx";
$outFile = "path/out.pdf";

$data = [
    'var1' => 'lorem',
    'var2' => 'ipsum',
    'img1' => 'path/image.jpg',
    't:tVar1' => [
        [
            'tVar1' => 'dolor',
            'tVar2' => 'sit',
        ],
        [
            'tVar1' => 'amet',
            'tVar2' => 'consectetur',
        ],
    ],
    'b:BLOCK' => [
        [
            'var1' => 'adipiscing',
            'var2' => 'elit',
        ],
        [
            'var1' => 'sed',
            'var2' => 'do',
        ],
        [
            'var1' => 'eiusmod',
            'var2' => 'tempor',
        ],
    ]
];

DocumentIO::process($templateFile, $outFile, $data);
```



##### Template.docx

This is a simple test template content with two variables **${var1}** and **${var2}**

With an image:

**${img1:300:200}**

With a table:

| Table variable 1 | Table variable 2 |
| ---------------- | ---------------- |
| **${tVar1}**     | **${tVar2}**     |

And a block:

**${BLOCK}**

This is a block content with two variables **${var1}** and **${var2}**

**${/BLOCK}**
