<?php
/**
 * Created by jacob.
 * Date: 2016/5/19 0019
 * Time: 下午 17:07
 */

namespace Jacobcyl\AliOSS;

use Dingo\Api\Contract\Transformer\Adapter;
use League\Flysystem\Adapter\AbstractAdapter;
use League\Flysystem\AdapterInterface;
use League\Flysystem\Config;
use League\Flysystem\Util;
use OSS\Core\MimeTypes;
use OSS\Core\OssException;
use OSS\OssClient;
use Log;
use Symfony\Component\Filesystem\Exception\FileNotFoundException;

class AliOssAdapter extends AbstractAdapter
{
    /**
     * @var Log debug Mode true|false
     */
    protected $debug;
    /**
     * @var array
     */
    protected static $resultMap = [
        'Body'           => 'raw_contents',
        'Content-Length' => 'size',
        'ContentType'    => 'mimetype',
        'Size'           => 'size',
        'StorageClass'   => 'storage_class',
    ];

    /**
     * @var array
     */
    protected static $metaOptions = [
        'CacheControl',
        'Expires',
        'ServerSideEncryption',
        'Metadata',
        'ACL',
        'ContentType',
        'ContentDisposition',
        'ContentLanguage',
        'ContentEncoding',
    ];

    protected static $metaMap = [
        'CacheControl'         => 'Cache-Control',
        'Expires'              => 'Expires',
        'ServerSideEncryption' => 'x-oss-server-side-encryption',
        'Metadata'             => 'x-oss-metadata-directive',
        'ACL'                  => 'x-oss-object-acl',
        'ContentType'          => 'Content-Type',
        'ContentDisposition'   => 'Content-Disposition',
        'ContentLanguage'      => 'response-content-language',
        'ContentEncoding'      => 'Content-Encoding',
    ];

    //Aliyun OSS Client OssClient
    protected $client;
    //bucket name
    protected $bucket;

    protected $endPoint;
    
    protected $cdnDomain;

    protected $ssl;

    protected $isCname;

    //配置
    protected $options = [
        'Multipart'   => 128
    ];

    // 常用文件类型的MIME映射表
    // 文件类型映射content-type
    protected static $mimeMap = [
        // 图片
        'jpg'  => 'image/jpeg',
        'jpeg' => 'image/jpeg',
        'png'  => 'image/png',
        'gif'  => 'image/gif',
        'webp' => 'image/webp',
        'svg'  => 'image/svg+xml',
        'bmp'  => 'image/bmp',
        'ico'  => 'image/x-icon',
        // 文档
        'pdf'  => 'application/pdf',
        'doc'  => 'application/msword',
        'docx' => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'xls'  => 'application/vnd.ms-excel',
        'xlsx' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'ppt'  => 'application/vnd.ms-powerpoint',
        'pptx' => 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'txt'  => 'text/plain',
        'csv'  => 'text/csv',
        // 音频/视频
        'mp3'  => 'audio/mpeg',
        'mp4'  => 'video/mp4',
        'webm' => 'video/webm',
        'ogg'  => 'audio/ogg',
        'avi'  => 'video/x-msvideo',
        // 压缩包
        'zip'  => 'application/zip',
        'rar'  => 'application/vnd.rar',
        '7z'   => 'application/x-7z-compressed',
        'tar'  => 'application/x-tar',
        'gz'   => 'application/gzip',
        // 网页与代码
        'html' => 'text/html',
        'htm'  => 'text/html',
        'css'  => 'text/css',
        'js'   => 'application/javascript',
        'json' => 'application/json',
        'xml'  => 'application/xml',
        // 安卓安装包
        'apk'  => 'application/vnd.android.package-archive',
        // 其他
        'xltx' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.template',
        'potx' => 'application/vnd.openxmlformats-officedocument.presentationml.template',
        'ppsx' => 'application/vnd.openxmlformats-officedocument.presentationml.slideshow',
        'sldx' => 'application/vnd.openxmlformats-officedocument.presentationml.slide',
        'dotx' => 'application/vnd.openxmlformats-officedocument.wordprocessingml.template',
        'xlam' => 'application/vnd.ms-excel.addin.macroEnabled.12',
        'xlsb' => 'application/vnd.ms-excel.sheet.binary.macroEnabled.12',
        'hqx' => 'application/mac-binhex40',
        'cpt' => 'application/mac-compactpro',
        'rtf' => 'text/rtf',
        'mif' => 'application/vnd.mif',
        'odc' => 'application/vnd.oasis.opendocument.chart',
        'odb' => 'application/vnd.oasis.opendocument.database',
        'odf' => 'application/vnd.oasis.opendocument.formula',
        'odg' => 'application/vnd.oasis.opendocument.graphics',
        'otg' => 'application/vnd.oasis.opendocument.graphics-template',
        'odi' => 'application/vnd.oasis.opendocument.image',
        'odp' => 'application/vnd.oasis.opendocument.presentation',
        'otp' => 'application/vnd.oasis.opendocument.presentation-template',
        'ods' => 'application/vnd.oasis.opendocument.spreadsheet',
        'ots' => 'application/vnd.oasis.opendocument.spreadsheet-template',
        'odt' => 'application/vnd.oasis.opendocument.text',
        'odm' => 'application/vnd.oasis.opendocument.text-master',
        'ott' => 'application/vnd.oasis.opendocument.text-template',
        'oth' => 'application/vnd.oasis.opendocument.text-web',
        'sxw' => 'application/vnd.sun.xml.writer',
        'stw' => 'application/vnd.sun.xml.writer.template',
        'sxc' => 'application/vnd.sun.xml.calc',
        'stc' => 'application/vnd.sun.xml.calc.template',
        'sxd' => 'application/vnd.sun.xml.draw',
        'std' => 'application/vnd.sun.xml.draw.template',
        'sxi' => 'application/vnd.sun.xml.impress',
        'sti' => 'application/vnd.sun.xml.impress.template',
        'sxg' => 'application/vnd.sun.xml.writer.global',
        'sxm' => 'application/vnd.sun.xml.math',
        'sis' => 'application/vnd.symbian.install',
        'wbxml' => 'application/vnd.wap.wbxml',
        'wmlc' => 'application/vnd.wap.wmlc',
        'wmlsc' => 'application/vnd.wap.wmlscriptc',
        'bcpio' => 'application/x-bcpio',
        'torrent' => 'application/x-bittorrent',
        'bz2' => 'application/x-bzip2',
        'vcd' => 'application/x-cdlink',
        'pgn' => 'application/x-chess-pgn',
        'cpio' => 'application/x-cpio',
        'csh' => 'application/x-csh',
        'dvi' => 'application/x-dvi',
        'spl' => 'application/x-futuresplash',
        'gtar' => 'application/x-gtar',
        'hdf' => 'application/x-hdf',
        'jar' => 'application/java-archive',
        'jnlp' => 'application/x-java-jnlp-file',
        'ksp' => 'application/x-kspread',
        'chrt' => 'application/x-kchart',
        'kil' => 'application/x-killustrator',
        'latex' => 'application/x-latex',
        'rpm' => 'application/x-rpm',
        'sh' => 'application/x-sh',
        'shar' => 'application/x-shar',
        'swf' => 'application/x-shockwave-flash',
        'sit' => 'application/x-stuffit',
        'sv4cpio' => 'application/x-sv4cpio',
        'sv4crc' => 'application/x-sv4crc',
        'tcl' => 'application/x-tcl',
        'tex' => 'application/x-tex',
        'man' => 'application/x-troff-man',
        'me' => 'application/x-troff-me',
        'ms' => 'application/x-troff-ms',
        'ustar' => 'application/x-ustar',
        'src' => 'application/x-wais-source',
        'm3u' => 'audio/x-mpegurl',
        'ra' => 'audio/x-pn-realaudio',
        'wav' => 'audio/x-wav',
        'wma' => 'audio/x-ms-wma',
        'wax' => 'audio/x-ms-wax',
        'pdb' => 'chemical/x-pdb',
        'xyz' => 'chemical/x-xyz',
        'ief' => 'image/ief',
        'wbmp' => 'image/vnd.wap.wbmp',
        'ras' => 'image/x-cmu-raster',
        'pnm' => 'image/x-portable-anymap',
        'pbm' => 'image/x-portable-bitmap',
        'pgm' => 'image/x-portable-graymap',
        'ppm' => 'image/x-portable-pixmap',
        'rgb' => 'image/x-rgb',
        'xbm' => 'image/x-xbitmap',
        'xpm' => 'image/x-xpixmap',
        'xwd' => 'image/x-xwindowdump',
        'rtx' => 'text/richtext',
        'tsv' => 'text/tab-separated-values',
        'jad' => 'text/vnd.sun.j2me.app-descriptor',
        'wml' => 'text/vnd.wap.wml',
        'wmls' => 'text/vnd.wap.wmlscript',
        'etx' => 'text/x-setext',
        'mxu' => 'video/vnd.mpegurl',
        'flv' => 'video/x-flv',
        'wm' => 'video/x-ms-wm',
        'wmv' => 'video/x-ms-wmv',
        'wmx' => 'video/x-ms-wmx',
        'wvx' => 'video/x-ms-wvx',
        'movie' => 'video/x-sgi-movie',
        'ice' => 'x-conference/x-cooltalk',
        '3gp' => 'video/3gpp',
        'ai' => 'application/postscript',
        'aif' => 'audio/x-aiff',
        'aifc' => 'audio/x-aiff',
        'aiff' => 'audio/x-aiff',
        'asc' => 'text/plain',
        'atom' => 'application/atom+xml',
        'au' => 'audio/basic',
        'bin' => 'application/octet-stream',
        'cdf' => 'application/x-netcdf',
        'cgm' => 'image/cgm',
        'class' => 'application/octet-stream',
        'dcr' => 'application/x-director',
        'dif' => 'video/x-dv',
        'dir' => 'application/x-director',
        'djv' => 'image/vnd.djvu',
        'djvu' => 'image/vnd.djvu',
        'dll' => 'application/octet-stream',
        'dmg' => 'application/octet-stream',
        'dms' => 'application/octet-stream',
        'dtd' => 'application/xml-dtd',
        'dv' => 'video/x-dv',
        'dxr' => 'application/x-director',
        'eps' => 'application/postscript',
        'exe' => 'application/octet-stream',
        'ez' => 'application/andrew-inset',
        'gram' => 'application/srgs',
        'grxml' => 'application/srgs+xml',
        'ics' => 'text/calendar',
        'ifb' => 'text/calendar',
        'iges' => 'model/iges',
        'igs' => 'model/iges',
        'jp2' => 'image/jp2',
        'jpe' => 'image/jpeg',
        'kar' => 'audio/midi',
        'lha' => 'application/octet-stream',
        'lzh' => 'application/octet-stream',
        'm4a' => 'audio/mp4a-latm',
        'm4p' => 'audio/mp4a-latm',
        'm4u' => 'video/vnd.mpegurl',
        'm4v' => 'video/x-m4v',
        'mac' => 'image/x-macpaint',
        'mathml' => 'application/mathml+xml',
        'mesh' => 'model/mesh',
        'mid' => 'audio/midi',
        'midi' => 'audio/midi',
        'mov' => 'video/quicktime',
        'mp2' => 'audio/mpeg',
        'mpe' => 'video/mpeg',
        'mpeg' => 'video/mpeg',
        'mpg' => 'video/mpeg',
        'mpga' => 'audio/mpeg',
        'msh' => 'model/mesh',
        'nc' => 'application/x-netcdf',
        'oda' => 'application/oda',
        'ogv' => 'video/ogv',
        'pct' => 'image/pict',
        'pic' => 'image/pict',
        'pict' => 'image/pict',
        'pnt' => 'image/x-macpaint',
        'pntg' => 'image/x-macpaint',
        'ps' => 'application/postscript',
        'qt' => 'video/quicktime',
        'qti' => 'image/x-quicktime',
        'qtif' => 'image/x-quicktime',
        'ram' => 'audio/x-pn-realaudio',
        'rdf' => 'application/rdf+xml',
        'rm' => 'application/vnd.rn-realmedia',
        'roff' => 'application/x-troff',
        'sgm' => 'text/sgml',
        'sgml' => 'text/sgml',
        'silo' => 'model/mesh',
        'skd' => 'application/x-koan',
        'skm' => 'application/x-koan',
        'skp' => 'application/x-koan',
        'skt' => 'application/x-koan',
        'smi' => 'application/smil',
        'smil' => 'application/smil',
        'snd' => 'audio/basic',
        'so' => 'application/octet-stream',
        't' => 'application/x-troff',
        'texi' => 'application/x-texinfo',
        'texinfo' => 'application/x-texinfo',
        'tif' => 'image/tiff',
        'tiff' => 'image/tiff',
        'tr' => 'application/x-troff',
        'vrml' => 'model/vrml',
        'vxml' => 'application/voicexml+xml',
        'wrl' => 'model/vrml',
        'xht' => 'application/xhtml+xml',
        'xhtml' => 'application/xhtml+xml',
        'xsl' => 'application/xml',
        'xslt' => 'application/xslt+xml',
        'xul' => 'application/vnd.mozilla.xul+xml',
    ];


    /**
     * AliOssAdapter constructor.
     *
     * @param OssClient $client
     * @param string    $bucket
     * @param string    $endPoint
     * @param bool      $ssl
     * @param bool      $isCname
     * @param bool      $debug
     * @param null      $prefix
     * @param array     $options
     */
    public function __construct(
        OssClient $client,
        $bucket,
        $endPoint,
        $ssl,
        $isCname = false,
        $debug = false,
        $cdnDomain,
        $prefix = null,
        array $options = []
    )
    {
        $this->debug = $debug;
        $this->client = $client;
        $this->bucket = $bucket;
        $this->setPathPrefix($prefix);
        $this->endPoint = $endPoint;
        $this->ssl = $ssl;
        $this->isCname = $isCname;
        $this->cdnDomain = $cdnDomain;
        $this->options = array_merge($this->options, $options);
    }

    /**
     * Get the OssClient bucket.
     *
     * @return string
     */
    public function getBucket()
    {
        return $this->bucket;
    }

    /**
     * Get the OSSClient instance.
     *
     * @return OssClient
     */
    public function getClient()
    {
        return $this->client;
    }

    /**
     * Write a new file.
     *
     * @param string $path
     * @param string $contents
     * @param Config $config Config object
     *
     * @return array|false false on failure file meta data on success
     */
    public function write($path, $contents, Config $config)
    {
        $object = $this->applyPathPrefix($path);
        $options = $this->getOptions($this->options, $config);

        if (! isset($options[OssClient::OSS_LENGTH])) {
            $options[OssClient::OSS_LENGTH] = Util::contentSize($contents);
        }
        $extension = pathinfo($contents, PATHINFO_EXTENSION);
        $content_type = self::getContentTypeByExtension($extension, MimeTypes::getMimetype($contents));
        $options[OssClient::OSS_CONTENT_TYPE] = $content_type ?: 'application/octet-stream';
        /*if (! isset($options[OssClient::OSS_CONTENT_TYPE])) {
            $options[OssClient::OSS_CONTENT_TYPE] = Util::guessMimeType($path, $contents);
        }*/
        try {
            $this->client->putObject($this->bucket, $object, $contents, $options);
        } catch (OssException $e) {
            $this->logErr(__FUNCTION__, $e);
            return false;
        }
        return $this->normalizeResponse($options, $path);
    }

    /**
     * Write a new file using a stream.
     *
     * @param string $path
     * @param resource $resource
     * @param Config $config Config object
     *
     * @return array|false false on failure file meta data on success
     */
    public function writeStream($path, $resource, Config $config)
    {
        $options = $this->getOptions($this->options, $config);
        $contents = stream_get_contents($resource);

        return $this->write($path, $contents, $config);
    }

    public function writeFile($path, $filePath, Config $config){
        $object = $this->applyPathPrefix($path);
        $options = $this->getOptions($this->options, $config);

        $options[OssClient::OSS_CHECK_MD5] = true;

        $extension = pathinfo($filePath, PATHINFO_EXTENSION);
        $content_type = self::getContentTypeByExtension($extension, MimeTypes::getMimetype($filePath));
        $options[OssClient::OSS_CONTENT_TYPE] = $content_type ?: 'application/octet-stream';
        /*if (! isset($options[OssClient::OSS_CONTENT_TYPE])) {
            $options[OssClient::OSS_CONTENT_TYPE] = Util::guessMimeType($path, $filePath);
        }*/
        try {
            $this->client->uploadFile($this->bucket, $object, $filePath, $options);
        } catch (OssException $e) {
            $this->logErr(__FUNCTION__, $e);
            return false;
        }
        return $this->normalizeResponse($options, $path);
    }

    /**
     * 通过文件后缀名获取对应的常见MIME类型
     * @param string $extension 文件后缀名 (如 'jpg', 'pdf')
     * @return string 对应的Content-Type，未找到则返回通用二进制流类型
     */
    public static function getContentTypeByExtension($extension, $default = 'application/octet-stream') {
        // 常用文件类型的MIME映射表
        // 获取后缀名的小写形式
        $extLower = strtolower($extension);
        // 返回对应的MIME类型，如果找不到则返回通用的二进制流类型
        return self::$mimeMap[$extLower] ?? $default;
    }

    /**
     * Update a file.
     *
     * @param string $path
     * @param string $contents
     * @param Config $config Config object
     *
     * @return array|false false on failure file meta data on success
     */
    public function update($path, $contents, Config $config)
    {
        if (! $config->has('visibility') && ! $config->has('ACL')) {
            $config->set(static::$metaMap['ACL'], $this->getObjectACL($path));
        }
        // $this->delete($path);
        return $this->write($path, $contents, $config);
    }

    /**
     * Update a file using a stream.
     *
     * @param string $path
     * @param resource $resource
     * @param Config $config Config object
     *
     * @return array|false false on failure file meta data on success
     */
    public function updateStream($path, $resource, Config $config)
    {
        $contents = stream_get_contents($resource);
        return $this->update($path, $contents, $config);
    }

    /**
     * {@inheritdoc}
     */
    public function rename($path, $newpath)
    {
        if (! $this->copy($path, $newpath)){
            return false;
        }

        return $this->delete($path);
    }

    /**
     * {@inheritdoc}
     */
    public function copy($path, $newpath)
    {
        $object = $this->applyPathPrefix($path);
        $newObject = $this->applyPathPrefix($newpath);
        try{
            $this->client->copyObject($this->bucket, $object, $this->bucket, $newObject);
        } catch (OssException $e) {
            $this->logErr(__FUNCTION__, $e);
            return false;
        }

        return true;
    }

    /**
     * {@inheritdoc}
     */
    public function delete($path)
    {
        $bucket = $this->bucket;
        $object = $this->applyPathPrefix($path);

        try{
            $this->client->deleteObject($bucket, $object);
        }catch (OssException $e) {
            $this->logErr(__FUNCTION__, $e);
            return false;
        }

        return ! $this->has($path);
    }

    /**
     * {@inheritdoc}
     */
    public function deleteDir($dirname)
    {
        $dirname = rtrim($this->applyPathPrefix($dirname), '/').'/';
        $dirObjects = $this->listDirObjects($dirname, true);

        if(count($dirObjects['objects']) > 0 ){

            foreach($dirObjects['objects'] as $object)
            {
                $objects[] = $object['Key'];
            }

            try {
                $this->client->deleteObjects($this->bucket, $objects);
            } catch (OssException $e) {
                $this->logErr(__FUNCTION__, $e);
                return false;
            }

        }

        try {
            $this->client->deleteObject($this->bucket, $dirname);
        } catch (OssException $e) {
            $this->logErr(__FUNCTION__, $e);
            return false;
        }

        return true;
    }

    /**
     * 列举文件夹内文件列表；可递归获取子文件夹；
     * @param string $dirname 目录
     * @param bool $recursive 是否递归
     * @return mixed
     * @throws OssException
     */
    public function listDirObjects($dirname = '', $recursive =  false)
    {
        $delimiter = '/';
        $nextMarker = '';
        $maxkeys = 1000;

        //存储结果
        $result = [];

        while(true){
            $options = [
                'delimiter' => $delimiter,
                'prefix'    => $dirname,
                'max-keys'  => $maxkeys,
                'marker'    => $nextMarker,
            ];

            try {
                $listObjectInfo = $this->client->listObjects($this->bucket, $options);
            } catch (OssException $e) {
                $this->logErr(__FUNCTION__, $e);
                // return false;
                throw $e;
            }

            $nextMarker = $listObjectInfo->getNextMarker(); // 得到nextMarker，从上一次listObjects读到的最后一个文件的下一个文件开始继续获取文件列表
            $objectList = $listObjectInfo->getObjectList(); // 文件列表
            $prefixList = $listObjectInfo->getPrefixList(); // 目录列表

            if (!empty($objectList)) {
                foreach ($objectList as $objectInfo) {

                    $object['Prefix']       = $dirname;
                    $object['Key']          = $objectInfo->getKey();
                    $object['LastModified'] = $objectInfo->getLastModified();
                    $object['eTag']         = $objectInfo->getETag();
                    $object['Type']         = $objectInfo->getType();
                    $object['Size']         = $objectInfo->getSize();
                    $object['StorageClass'] = $objectInfo->getStorageClass();

                    $result['objects'][] = $object;
                }
            }else{
                $result["objects"] = [];
            }

            if (!empty($prefixList)) {
                foreach ($prefixList as $prefixInfo) {
                    $result['prefix'][] = $prefixInfo->getPrefix();
                }
            }else{
                $result['prefix'] = [];
            }

            //递归查询子目录所有文件
            if($recursive){
                foreach( $result['prefix'] as $pfix){
                    $next  =  $this->listDirObjects($pfix , $recursive);
                    $result["objects"] = array_merge($result['objects'], $next["objects"]);
                }
            }

            //没有更多结果了
            if ($nextMarker === '') {
                break;
            }
        }

        return $result;
    }

    /**
     * {@inheritdoc}
     */
    public function createDir($dirname, Config $config)
    {
        $object = $this->applyPathPrefix($dirname);
        $options = $this->getOptionsFromConfig($config);

        try {
            $this->client->createObjectDir($this->bucket, $object, $options);
        } catch (OssException $e) {
            $this->logErr(__FUNCTION__, $e);
            return false;
        }

        return ['path' => $dirname, 'type' => 'dir'];
    }

    /**
     * {@inheritdoc}
     */
    public function setVisibility($path, $visibility)
    {
        $object = $this->applyPathPrefix($path);
        $acl = ( $visibility === AdapterInterface::VISIBILITY_PUBLIC ) ? OssClient::OSS_ACL_TYPE_PUBLIC_READ : OssClient::OSS_ACL_TYPE_PRIVATE;

        $this->client->putObjectAcl($this->bucket, $object, $acl);

        return compact('visibility');
    }

    /**
     * {@inheritdoc}
     */
    public function has($path)
    {
        $object = $this->applyPathPrefix($path);

        return $this->client->doesObjectExist($this->bucket, $object);
    }

    /**
     * {@inheritdoc}
     */
    public function read($path)
    {
        $result = $this->readObject($path);
        $result['contents'] = (string) $result['raw_contents'];
        unset($result['raw_contents']);
        return $result;
    }

    /**
     * {@inheritdoc}
     */
    public function readStream($path)
    {
        $result = $this->readObject($path);
        $result['stream'] = $result['raw_contents'];
        rewind($result['stream']);
        // Ensure the EntityBody object destruction doesn't close the stream
        $result['raw_contents']->detachStream();
        unset($result['raw_contents']);

        return $result;
    }

    /**
     * Read an object from the OssClient.
     *
     * @param string $path
     *
     * @return array
     */
    protected function readObject($path)
    {
        $object = $this->applyPathPrefix($path);

        $result['Body'] = $this->client->getObject($this->bucket, $object);
        $result = array_merge($result, ['type' => 'file']);
        return $this->normalizeResponse($result, $path);
    }

    /**
     * {@inheritdoc}
     */
    public function listContents($directory = '', $recursive = false)
    {
        $dirObjects = $this->listDirObjects($directory, true);
        $contents = $dirObjects["objects"];

        $result = array_map([$this, 'normalizeResponse'], $contents);
        $result = array_filter($result, function ($value) {
            return $value['path'] !== false;
        });

        return Util::emulateDirectories($result);
    }

    /**
     * {@inheritdoc}
     */
    public function getMetadata($path)
    {
        $object = $this->applyPathPrefix($path);

        try {
            $objectMeta = $this->client->getObjectMeta($this->bucket, $object);
        } catch (OssException $e) {
            $this->logErr(__FUNCTION__, $e);
            return false;
        }

        return $objectMeta;
    }

    /**
     * {@inheritdoc}
     */
    public function getSize($path)
    {
        $object = $this->getMetadata($path);
        $object['size'] = $object['content-length'];
        return $object;
    }

    /**
     * {@inheritdoc}
     */
    public function getMimetype($path)
    {
        if( $object = $this->getMetadata($path))
            $object['mimetype'] = $object['content-type'];
        return $object;
    }

    /**
     * {@inheritdoc}
     */
    public function getTimestamp($path)
    {
        if( $object = $this->getMetadata($path))
            $object['timestamp'] = strtotime( $object['last-modified'] );
        return $object;
    }

    /**
     * {@inheritdoc}
     */
    public function getVisibility($path)
    {
        $object = $this->applyPathPrefix($path);
        try {
            $acl = $this->client->getObjectAcl($this->bucket, $object);
        } catch (OssException $e) {
            $this->logErr(__FUNCTION__, $e);
            return false;
        }
        
        if ($acl == OssClient::OSS_ACL_TYPE_PUBLIC_READ ){
            $res['visibility'] = AdapterInterface::VISIBILITY_PUBLIC;
        }else{
            $res['visibility'] = AdapterInterface::VISIBILITY_PRIVATE;
        }

        return $res;
    }


    /**
     * @param $path
     *
     * @return string
     */
    public function getUrl( $path )
    {
        if (!$this->has($path)) throw new FileNotFoundException($filePath.' not found');
        return ( $this->ssl ? 'https://' : 'http://' ) . ( $this->isCname ? ( $this->cdnDomain == '' ? $this->endPoint : $this->cdnDomain ) : $this->bucket . '.' . $this->endPoint ) . '/' . ltrim($path, '/');
    }

    /**
     * The the ACL visibility.
     *
     * @param string $path
     *
     * @return string
     */
    protected function getObjectACL($path)
    {
        $metadata = $this->getVisibility($path);

        return $metadata['visibility'] === AdapterInterface::VISIBILITY_PUBLIC ? OssClient::OSS_ACL_TYPE_PUBLIC_READ : OssClient::OSS_ACL_TYPE_PRIVATE;
    }

    /**
     * Normalize a result from OSS.
     *
     * @param array  $object
     * @param string $path
     *
     * @return array file metadata
     */
    protected function normalizeResponse(array $object, $path = null)
    {
        $result = ['path' => $path ?: $this->removePathPrefix(isset($object['Key']) ? $object['Key'] : $object['Prefix'])];
        $result['dirname'] = Util::dirname($result['path']);

        if (isset($object['LastModified'])) {
            $result['timestamp'] = strtotime($object['LastModified']);
        }

        if (substr($result['path'], -1) === '/') {
            $result['type'] = 'dir';
            $result['path'] = rtrim($result['path'], '/');

            return $result;
        }
        
        $result = array_merge($result, Util::map($object, static::$resultMap), ['type' => 'file']);

        return $result;
    }

    /**
     * Get options for a OSS call. done
     *
     * @param array  $options
     *
     * @return array OSS options
     */
    protected function getOptions(array $options = [], Config $config = null)
    {
        $options = array_merge($this->options, $options);

        if ($config) {
            $options = array_merge($options, $this->getOptionsFromConfig($config));
        }

        return array(OssClient::OSS_HEADERS => $options);
    }

    /**
     * Retrieve options from a Config instance. done
     *
     * @param Config $config
     *
     * @return array
     */
    protected function getOptionsFromConfig(Config $config)
    {
        $options = [];

        foreach (static::$metaOptions as $option) {
            if (! $config->has($option)) {
                continue;
            }
            $options[static::$metaMap[$option]] = $config->get($option);
        }

        if ($visibility = $config->get('visibility')) {
            // For local reference
            // $options['visibility'] = $visibility;
            // For external reference
            $options['x-oss-object-acl'] = $visibility === AdapterInterface::VISIBILITY_PUBLIC ? OssClient::OSS_ACL_TYPE_PUBLIC_READ : OssClient::OSS_ACL_TYPE_PRIVATE;
        }

        if ($mimetype = $config->get('mimetype')) {
            // For local reference
            // $options['mimetype'] = $mimetype;
            // For external reference
            $options['Content-Type'] = $mimetype;
        }

        return $options;
    }

    /**
     * @param $fun string function name : __FUNCTION__
     * @param $e
     */
    protected function logErr($fun, $e){
        if( $this->debug ){
            Log::error($fun . ": FAILED");
            Log::error($e->getMessage());
        }
    }
}
