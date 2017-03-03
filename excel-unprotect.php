#!/usr/bin/env php
<?php

if (count($argv) < 2) {
  print "Usage: excel-unprotect file1.xlsx file2.xlsx [...]\n";
  exit(1);
}

foreach (array_slice($argv, 1) as $file) {
  if (!file_exists($file)) {
    echo "Skipping non-existing file $file\n";
    continue;
  }

  $pathinfo = pathinfo($file);
  if (empty($pathinfo['extension']) || $pathinfo['extension'] !== 'xlsx') {
    echo "Only XLSX files are supported. Skipping $file\n";
    continue;
  }

  // Create a temporary folder to unpack the file into.
  $dir = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'excel-unprotect-' . uniqid();
  if (!mkdir($dir)) {
    echo "Could not create temporary directory. Aborting.\n";
    exit(2);
  }
  $dir = realpath($dir);

  // Extract the XLSX file.
  $zip = new ZipArchive();
  if ($zip->open($file) !== TRUE) {
    echo "Could not open $file. Skipping.\n";
    continue;
  }

  $zip->extractTo($dir);
  $zip->close();

  // Loop over all worksheets and remove the protection.
  $worksheet_dir = implode(DIRECTORY_SEPARATOR, [$dir, 'xl', 'worksheets']);
  $document_unprotected = FALSE;
  foreach (glob($worksheet_dir . DIRECTORY_SEPARATOR . '*.xml') as $worksheet) {
    $document = new DOMDocument();
    $document->load($worksheet);
    $document_element = $document->documentElement;

    $document_changed = FALSE;
    foreach ($document_element->getElementsByTagName('sheetProtection') as $element) {
      $document_element->removeChild($element);
      $document_changed = $document_unprotected = TRUE;
    }

    if ($document_changed) {
      $document->save($worksheet);
    }
  }

  // Save an unprotected copy of the file.
  if ($document_unprotected) {
    $unprotected_file = $pathinfo['dirname'] . DIRECTORY_SEPARATOR . $pathinfo['filename'] . '-unprotected.' . $pathinfo['extension'];
    if (file_exists($unprotected_file)) {
      echo "Unprotected file $unprotected_file already exists. Skipping.\n";
      continue;
    }

    $zip = new ZipArchive();
    if ($zip->open($unprotected_file, ZIPARCHIVE::CREATE) !== TRUE) {
      echo "Can't open file $unprotected_file for writing. Skipping.\n";
      continue;
    }

    foreach (new RecursiveIteratorIterator(new RecursiveDirectoryIterator($dir), RecursiveIteratorIterator::LEAVES_ONLY) as $file) {
      if (!$file->isDir()) {
        $file_path = $file->getRealPath();
        $zip->addFile($file_path, substr($file_path, strlen($dir) + 1));
      }
    }
    $zip->close();
  } else {
    echo "No sheet protection found in $file. Skipping.\n";
  }
}
