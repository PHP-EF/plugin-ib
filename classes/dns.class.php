<?php
// ibPortal DNS Class
// This class provides a way to perform DNS queries using DNS over HTTPS (DoH).
class ibDNS extends ibPortal {

    public function __construct() {
		parent::__construct();
	}

    public string $dohEndpoint = 'https://cloudflare-dns.com/dns-query';

    public function dohQuery(string $domain, int $type = 1, string $dohDomain = null): array {
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

        return $this->parseResponse($response->body);
    }

    protected function buildQuery(string $domain, int $type): string {
        $id = random_int(0, 0xffff);
        $flags = 0x0100;
        $qdcount = 1;

        $header = pack('nnnnnn', $id, $flags, $qdcount, 0, 0, 0);

        $qname = '';
        foreach (explode('.', $domain) as $label) {
            $qname .= chr(strlen($label)) . $label;
        }
        $qname .= "\0";

        $question = $qname . pack('nn', $type, 1); // QTYPE, QCLASS

        return $header . $question;
    }

    protected function parseResponse(string $data): array {
        $flags = unpack('n', substr($data, 2, 2))[1];
        $rcode = $flags & 0x000F;
    
        if ($rcode === 3) {
            return ['error' => 'NXDOMAIN: Domain does not exist'];
        }

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

        for ($i = 0; $i < $ancount; $i++) {
            
            $byte = ord($data[$offset]);
            $offset += ($byte & 0xC0) === 0xC0 ? 2 : strlen($this->readName($data, $offset)) + 2;

            $type = unpack('n', substr($data, $offset, 2))[1];
            $offset += 2 + 2 + 4; // type + class + TTL
            $rdlength = unpack('n', substr($data, $offset, 2))[1];
            $offset += 2;

            $rdata = substr($data, $offset, $rdlength);

            switch ($type) {
                case 1: // A
                    $records[] = ['type' => 'A', 'value' => inet_ntop($rdata)];
                    break;
                case 2: // NS
                    $ns = $this->readName($data, $offset);
                    $records[] = ['type' => 'NS', 'value' => $ns];
                    break;
                case 5: // CNAME
                    $cname = $this->readName($data, $offset);
                    $records[] = ['type' => 'CNAME', 'value' => $cname];
                    break;
                case 6: // SOA
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
                        'minimum' => $minimum
                    ];
                    break;
                case 12: // PTR
                    $ptr = $this->readName($data, $offset);
                    $records[] = ['type' => 'PTR', 'value' => $ptr];
                    break;
                case 15: // MX
                    $priority = unpack('n', substr($rdata, 0, 2))[1];
                    $exchange = $this->readName($data, $offset + 2);
                    $records[] = ['type' => 'MX', 'priority' => $priority, 'exchange' => $exchange];
                    break;
                case 16: // TXT
                    $txtLen = ord($rdata[0]);
                    $txt = substr($rdata, 1, $txtLen);
                    $records[] = ['type' => 'TXT', 'value' => $txt];
                    break;
                case 28: // AAAA
                    $records[] = ['type' => 'AAAA', 'value' => inet_ntop($rdata)];
                    break;
                case 43: // DS
                    $keyTag = unpack('n', substr($rdata, 0, 2))[1];
                    $algorithm = ord($rdata[2]);
                    $digestType = ord($rdata[3]);
                    $digest = substr($rdata, 4);
                    $records[] = [
                        'type' => 'DS',
                        'keyTag' => $keyTag,
                        'algorithm' => $algorithm,
                        'digestType' => $digestType,
                        'digest' => bin2hex($digest)
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
        