<?php
use RemotelyLiving\PHPDNS\Resolvers;
use RemotelyLiving\PHPDNS\Entities;

class DNSToolbox {

    public function a($hostname,$sourceserver) {
        $nameserver = new Entities\Hostname($sourceserver);
        $resolver = new Resolvers\Dig(null,null,$nameserver);
        $records = $resolver->getARecords($hostname);
        return $records;
    }

    public function aaaa($hostname,$sourceserver) {
        $nameserver = new Entities\Hostname($sourceserver);
        $resolver = new Resolvers\Dig(null,null,$nameserver);
        $records = $resolver->getAAAARecords($hostname);
        return $records;
    }

    public function cname($hostname,$sourceserver) {
        $nameserver = new Entities\Hostname($sourceserver);
        $resolver = new Resolvers\Dig(null,null,$nameserver);
        $records = $resolver->getCNAMERecords($hostname);
        return $records;
    }

    public function all($hostname,$sourceserver) {
        $nameserver = new Entities\Hostname($sourceserver);
        $resolver = new Resolvers\Dig(null,null,$nameserver);
        $records = $resolver->getRecords($hostname);
        return $records;
    }

    public function dmarc($hostname,$sourceserver) {
        $dmarc = "_dmarc.";
        $dmarchostname = $dmarc . $hostname;
        $nameserver = new Entities\Hostname($sourceserver);
        $resolver = new Resolvers\Dig(null,null,$nameserver);
        $records = $resolver->getTXTRecords($dmarchostname);
        return $records;
    }

    public function txt($hostname,$sourceserver) {
        $nameserver = new Entities\Hostname($sourceserver);
        $resolver = new Resolvers\Dig(null,null,$nameserver);
        $records = $resolver->getTxtRecords($hostname);
        return $records;
    }

    public function mx($hostname,$sourceserver) {
        $nameserver = new Entities\Hostname($sourceserver);
        $resolver = new Resolvers\Dig(null,null,$nameserver);
        $records = $resolver->getMXRecords($hostname);
        return $records;
    }

    public function ns($domain,$sourceserver) {
        $nameserver = new Entities\Hostname($sourceserver);
        $resolver = new Resolvers\Dig(null,null,$nameserver);
        $dnsType = new Entities\DNSRecordType("NS");
        $records = $resolver->getRecords($domain,$dnsType);
        return $records;
    }

    public function soa($hostname,$sourceserver) {
        $nameserver = new Entities\Hostname($sourceserver);
        $resolver = new Resolvers\Dig(null,null,$nameserver);
        $dnsType = new Entities\DNSRecordType("SOA");
        $records = $resolver->getRecords($hostname,$dnsType);
        return $records;
    }

    public function reverse($ip,$source) {
        if((bool)ip2long($ip)){
            $response = gethostbyaddr($ip);
            return [array(
                'ip' => $ip,
                'hostname' => $response
            )];
        } else {
            return array(
                'Status' => 'Error',
                'Message' => 'Invalid IP Address'
            );
        }
    }

    public function port($hostname,$port) {
        if (count($port) == 0) {
            $port = [22, 25, 53, 80 , 443, 445, 3389];
        }
        $portList = [];
        for ($i = 0; $i < count($port); $i++) {
            $fp = fsockopen($hostname, $port[$i], $errno, $errstr, 5);
            if ($fp) {
                $result = 'Open';
                fclose($fp);
            } else {
                $result = 'Closed';
            }
            $portList[] = array(
                'hostname' => $hostname,
                'port' => $port[$i],
                'result' => $result
            );
        }
        return $portList;
    }

    // ** DNS over HTTPS (DoH) Implementation ** //

    protected const RECORD_TYPES = [
        'a' => 1,
        'ns' => 2,
        'cname' => 5,
        'soa' => 6,
        'ptr' => 12,
        'mx' => 15,
        'txt' => 16,
        'aaaa' => 28,
        'ds' => 43,
    ];

    public string $dohEndpoint = 'https://cloudflare-dns.com/dns-query';

    public function dohQuery(string $domain, $type, string $dohDomain = null): array {
        if ($dohDomain != null) {
            $this->dohEndpoint = 'https://' . $dohDomain . '/dns-query';
        }
        $query = $this->buildQuery($domain, $type);
        $headers = [
            'Content-Type' => 'application/dns-message',
            'Accept' => 'application/dns-message'
        ];
        
        $response = Requests::post($this->dohEndpoint, $headers, $query);

        if ($response->status_code !== 200) {
            throw new Exception("DoH query failed with status: " . $response->status_code);
        }

        return $this->parseResponse($domain, $response->body);
    }


    protected function convertRecordType(string|int $value): string|int
        {
            if (is_string($value)) {
                if (isset(self::RECORD_TYPES[$value])) {
                    return self::RECORD_TYPES[$value];
                }
                throw new Exception("Unsupported DNS record type: " . $value);
            } elseif (is_int($value)) {
                $type = array_search($value, self::RECORD_TYPES, true);
                if ($type !== false) {
                    return $type;
                }
                throw new Exception("Unsupported DNS record code: " . $value);
            }
            throw new Exception("Invalid input type. Must be string or int.");
    }

    protected function buildQuery(string $domain, string $type): string {
        $typeInt = $this->convertRecordType($type);
        $id = random_int(0, 0xffff);
        $flags = 0x0100;
        $qdcount = 1;

        $header = pack('nnnnnn', $id, $flags, $qdcount, 0, 0, 0);

        $qname = '';
        foreach (explode('.', $domain) as $label) {
            $qname .= chr(strlen($label)) . $label;
        }
        $qname .= "\0";

        $question = $qname . pack('nn', $typeInt, 1); // QTYPE, QCLASS

        return $header . $question;
    }

    protected function parseResponse(string $hostname, string $data): array {
        
        $flags = unpack('n', substr($data, 2, 2))[1];
        $rcode = $flags & 0x000F;
    
        $offset = 12;
        $qdcount = unpack('n', substr($data, 4, 2))[1];
        $ancount = unpack('n', substr($data, 6, 2))[1];

        for ($i = 0; $i < $qdcount; $i++) {
            while (ord($data[$offset]) !== 0) {
                $offset += ord($data[$offset]) + 1;
            }
            $offset += 5;
        }

        $records = [];

        if ($rcode === 3) {
            return array(['data' => 'NXDOMAIN', 'IPAddress' => 'NXDOMAIN', 'hostname' => $hostname]);
        }

        for ($i = 0; $i < $ancount; $i++) {
            
            $byte = ord($data[$offset]);
            $offset += ($byte & 0xC0) === 0xC0 ? 2 : strlen($this->readName($data, $offset)) + 2;

            $typeint = unpack('n', substr($data, $offset, 2))[1];
            $type = $this->convertRecordType($typeint);
            $offset += 2 + 2 + 4; // type + class + TTL
            $rdlength = unpack('n', substr($data, $offset, 2))[1];
            $offset += 2;

            // calculate ttl
            $ttl = unpack('N', substr($data, $offset - 6, 4))[1];
            if ($ttl < 0) {
                $ttl = 0;
            }

            $rdata = substr($data, $offset, $rdlength);

            switch ($type) {
                case 'a': // A
                    $records[] = ['type' => 'A', 'IPAddress' => inet_ntop($rdata), 'hostname' => $hostname, 'TTL' => $ttl];
                    break;
                case 'ns': // NS
                    $ns = $this->readName($data, $offset);
                    $records[] = ['type' => 'NS', 'data' => $ns, 'hostname' => $hostname, 'TTL' => $ttl];
                    break;
                case 'cname': // CNAME
                    $cname = $this->readName($data, $offset);
                    $records[] = ['type' => 'CNAME', 'data' => $cname, 'hostname' => $hostname, 'TTL' => $ttl];
                    break;
                case 'soa': // SOA
                    $mname = $this->readName($data, $offset);
                    $rname = $this->readName($data, $offset + strlen($mname) + 1);
                    $serial = unpack('N', substr($rdata, 0, 4))[1];
                    $refresh = unpack('N', substr($rdata, 4, 4))[1];
                    $retry = unpack('N', substr($rdata, 8, 4))[1];
                    $expire = unpack('N', substr($rdata, 12, 4))[1];
                    $minimum = unpack('N', substr($rdata, 16, 4))[1];
                    $records[] = [
                        'type' => 'SOA',
                        'mname' => $mname,
                        'rname' => $rname,
                        'serial' => $serial,
                        'refresh' => $refresh,
                        'retry' => $retry,
                        'expire' => $expire,
                        'minimum' => $minimum,
                        'hostname' => $hostname,
                        'TTL' => $ttl
                    ];
                    break;
                case 'ptr': // PTR
                    $ptr = $this->readName($data, $offset);
                    $records[] = ['type' => 'PTR', 'data' => $ptr, 'hostname' => $hostname, 'TTL' => $ttl];
                    break;
                case 'mx': // MX
                    $priority = unpack('n', substr($rdata, 0, 2))[1];
                    $exchange = $this->readName($data, $offset + 2);
                    $records[] = ['type' => 'MX', 'priority' => $priority, 'exchange' => $exchange, 'hostname' => $hostname, 'TTL' => $ttl];
                    break;
                case 'txt': // TXT
                    $txtLen = ord($rdata[0]);
                    $txt = substr($rdata, 1, $txtLen);
                    $records[] = ['type' => 'TXT', 'data' => $txt];
                    break;
                case 'aaaa': // AAAA
                    $records[] = ['type' => 'AAAA', 'IPAddress' => inet_ntop($rdata), 'hostname' => $hostname, 'TTL' => $ttl];
                    break;
                case 'ds': // DS
                    $keyTag = unpack('n', substr($rdata, 0, 2))[1];
                    $algorithm = ord($rdata[2]);
                    $digestType = ord($rdata[3]);
                    $digest = substr($rdata, 4);
                    $records[] = [
                        'type' => 'DS',
                        'keyTag' => $keyTag,
                        'algorithm' => $algorithm,
                        'digestType' => $digestType,
                        'digest' => bin2hex($digest),
                        'hostname' => $hostname,
                        'TTL' => $ttl
                    ];
                    break;
                default:
                    $records[] = ['type' => $type, 'raw' => bin2hex($rdata)];
            }

            $offset += $rdlength;
        }

        return $records;
    }

    protected function readName(string $data, int $offset): string {
        $labels = [];
        while (true) {
            $len = ord($data[$offset]);
            if ($len === 0) {
                $offset++;
                break;
            }
            if (($len & 0xC0) === 0xC0) {
                $pointer = unpack('n', substr($data, $offset, 2))[1] & 0x3FFF;
                $labels[] = $this->readName($data, $pointer);
                $offset += 2;
                break;
            } else {
                $labels[] = substr($data, $offset + 1, $len);
                $offset += $len + 1;
            }
        }
        return implode('.', $labels);
    }
}