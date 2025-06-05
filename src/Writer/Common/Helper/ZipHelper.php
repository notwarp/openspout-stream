<?php

declare(strict_types=1);

namespace OpenSpout\Writer\Common\Helper;

use RecursiveDirectoryIterator;
use RecursiveIteratorIterator;
use SplFileInfo;
use ZipArchive;
use ZipStream\CompressionMethod;
use ZipStream\Exception\FileNotFoundException;
use ZipStream\Exception\FileNotReadableException;
use ZipStream\Exception\OverflowException;
use ZipStream\ZipStream;

/**
 * @internal
 */
final class ZipHelper
{
    public const ZIP_EXTENSION = '.zip';

    /**
     * Controls what to do when trying to add an existing file.
     */
    public const EXISTING_FILES_SKIP = 'skip';
    public const EXISTING_FILES_OVERWRITE = 'overwrite';
    public null|ZipStream $zip;

    public array $already_putted_files = [];

    public function __construct(null|ZipStream $zip)
    {
        $this->zip = $zip;
    }

    /**
     * Returns a new ZipArchive instance pointing at the given path.
     *
     * @param string $tmpFolderPath Path of the temp folder where the zip file will be created
     */
    public function createZip(string $tmpFolderPath): ZipArchive|ZipStream
    {
        if (null === $this->zip) {
            $zip = new ZipArchive();
            $zipFilePath = $tmpFolderPath.self::ZIP_EXTENSION;

            $zip->open($zipFilePath, ZipArchive::CREATE | ZipArchive::OVERWRITE);

            return $zip;
        }

        return $this->zip;
    }

    /**
     * @param ZipArchive|ZipStream $zip An opened zip archive object
     *
     * @return string Path where the zip file of the given folder will be created
     */
    public function getZipFilePath(ZipArchive|ZipStream $zip): string
    {
        if ($zip instanceof ZipArchive) {
            return $zip->filename;
        }

        return '';
    }

    /**
     * Adds the given file, located under the given root folder to the archive.
     * The file will be compressed.
     *
     * Example of use:
     *   addFileToArchive($zip, '/tmp/xlsx/foo', 'bar/baz.xml');
     *   => will add the file located at '/tmp/xlsx/foo/bar/baz.xml' in the archive, but only as 'bar/baz.xml'
     *
     * @param ZipArchive|ZipStream $zip              An opened zip archive object
     * @param string               $rootFolderPath   path of the root folder that will be ignored in the archive tree
     * @param string               $localFilePath    Path of the file to be added, under the root folder
     * @param string               $existingFileMode Controls what to do when trying to add an existing file
     */
    public function addFileToArchive(ZipArchive|ZipStream $zip, string $rootFolderPath, string $localFilePath, string $existingFileMode = self::EXISTING_FILES_OVERWRITE): void
    {
        $this->addFileToArchiveWithCompressionMethod(
            $zip,
            $rootFolderPath,
            $localFilePath,
            $existingFileMode,
            ($zip instanceof ZipArchive) ? ZipArchive::CM_DEFAULT : CompressionMethod::DEFLATE
        );
        $this->already_putted_files[] = $localFilePath;
    }

    /**
     * Adds the given file, located under the given root folder to the archive.
     * The file will NOT be compressed.
     *
     * Example of use:
     *   addUncompressedFileToArchive($zip, '/tmp/xlsx/foo', 'bar/baz.xml');
     *   => will add the file located at '/tmp/xlsx/foo/bar/baz.xml' in the archive, but only as 'bar/baz.xml'
     *
     * @param ZipArchive|ZipStream $zip              An opened zip archive object
     * @param string               $rootFolderPath   path of the root folder that will be ignored in the archive tree
     * @param string               $localFilePath    Path of the file to be added, under the root folder
     * @param string               $existingFileMode Controls what to do when trying to add an existing file
     */
    public function addUncompressedFileToArchive(ZipArchive|ZipStream $zip, string $rootFolderPath, string $localFilePath, string $existingFileMode = self::EXISTING_FILES_OVERWRITE): void
    {
        $this->addFileToArchiveWithCompressionMethod(
            $zip,
            $rootFolderPath,
            $localFilePath,
            $existingFileMode,
            ($zip instanceof ZipArchive) ? ZipArchive::CM_STORE : CompressionMethod::STORE
        );
    }

    /**
     * @param ZipArchive|ZipStream $zip              An opened zip archive object
     * @param string               $folderPath       Path to the folder to be zipped
     * @param string               $existingFileMode Controls what to do when trying to add an existing file
     */
    public function addFolderToArchive(ZipArchive|ZipStream $zip, string $folderPath, string $existingFileMode = self::EXISTING_FILES_OVERWRITE): void
    {
        $folderRealPath = $this->getNormalizedRealPath($folderPath).'/';
        $itemIterator = new RecursiveIteratorIterator(
            new RecursiveDirectoryIterator($folderPath, RecursiveDirectoryIterator::SKIP_DOTS),
            RecursiveIteratorIterator::SELF_FIRST
        );

        foreach ($itemIterator as $itemInfo) {
            \assert($itemInfo instanceof SplFileInfo);
            $itemRealPath = $this->getNormalizedRealPath($itemInfo->getPathname());
            $itemLocalPath = str_replace($folderRealPath, '', $itemRealPath);

            if ($itemInfo->isFile() && !$this->shouldSkipFile($zip, $itemLocalPath, $existingFileMode)
                && !\in_array($itemLocalPath, $this->already_putted_files, true)
            ) {
                if ($zip instanceof ZipArchive) {
                    $zip->addFile($itemRealPath, $itemLocalPath);
                } else {
                    $zip->addFileFromPath(
                        fileName: $itemLocalPath,
                        path: $itemRealPath
                    );
                }
                $this->already_putted_files[] = $itemLocalPath;
            }
        }
    }

    /**
     * Closes the archive and copies it into the given stream.
     *
     * @param ZipArchive|ZipStream $zip           An opened zip archive object
     * @param resource             $streamPointer Pointer to the stream to copy the zip
     *
     * @throws OverflowException
     */
    public function closeArchiveAndCopyToStream(ZipArchive|ZipStream $zip, $streamPointer): void
    {
        if ($zip instanceof ZipArchive) {
            $zipFilePath = $zip->filename;
            $zip->close();

            $this->copyZipToStream($zipFilePath, $streamPointer);
        } else {
            $zip->finish();
        }
    }

    /**
     * Adds the given file, located under the given root folder to the archive.
     * The file will NOT be compressed.
     *
     * Example of use:
     *   addUncompressedFileToArchive($zip, '/tmp/xlsx/foo', 'bar/baz.xml');
     *   => will add the file located at '/tmp/xlsx/foo/bar/baz.xml' in the archive, but only as 'bar/baz.xml'
     *
     * @param ZipArchive|ZipStream  $zip               An opened zip archive object
     * @param string                $rootFolderPath    path of the root folder that will be ignored in the archive tree
     * @param string                $localFilePath     Path of the file to be added, under the root folder
     * @param string                $existingFileMode  Controls what to do when trying to add an existing file
     * @param CompressionMethod|int $compressionMethod The compression method
     *
     * @throws FileNotFoundException
     * @throws FileNotReadableException
     */
    private function addFileToArchiveWithCompressionMethod(ZipArchive|ZipStream $zip, string $rootFolderPath, string $localFilePath, string $existingFileMode, CompressionMethod|int $compressionMethod): void
    {
        $normalizedLocalFilePath = str_replace('\\', '/', $localFilePath);
        if (!$this->shouldSkipFile($zip, $normalizedLocalFilePath, $existingFileMode)) {
            $normalizedFullFilePath = $this->getNormalizedRealPath($rootFolderPath.'/'.$normalizedLocalFilePath);
            if ($zip instanceof ZipArchive) {
                $zip->addFile($normalizedFullFilePath, $normalizedLocalFilePath);
                $zip->setCompressionName($normalizedLocalFilePath, $compressionMethod);
            } else {
                $zip->addFileFromPath(
                    fileName: $normalizedLocalFilePath,
                    path: $normalizedFullFilePath,
                    compressionMethod: CompressionMethod::STORE
                );
            }
        }
    }

    /**
     * @return bool Whether the file should be added to the archive or skipped
     */
    private function shouldSkipFile(ZipArchive|ZipStream $zip, string $itemLocalPath, string $existingFileMode): bool
    {
        // Skip files if:
        //   - EXISTING_FILES_SKIP mode chosen
        //   - File already exists in the archive
        return self::EXISTING_FILES_SKIP === $existingFileMode
            && false !== (
                $zip instanceof ZipArchive ?
                $zip->locateName($itemLocalPath) :
                false
            );
    }

    /**
     * Returns canonicalized absolute pathname, containing only forward slashes.
     *
     * @param string $path Path to normalize
     *
     * @return string Normalized and canonicalized path
     */
    private function getNormalizedRealPath(string $path): string
    {
        $realPath = realpath($path);
        \assert(false !== $realPath);

        return str_replace(\DIRECTORY_SEPARATOR, '/', $realPath);
    }

    /**
     * Streams the contents of the zip file into the given stream.
     *
     * @param string   $zipFilePath Path of the zip file
     * @param resource $pointer     Pointer to the stream to copy the zip
     */
    private function copyZipToStream(string $zipFilePath, $pointer): void
    {
        $zipFilePointer = fopen($zipFilePath, 'r');
        \assert(false !== $zipFilePointer);
        stream_copy_to_stream($zipFilePointer, $pointer);
        fclose($zipFilePointer);
    }
}
