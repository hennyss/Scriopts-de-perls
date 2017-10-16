#!/usr/bin/perl
# aptitude install libpdf-report-perl

use lib '/home/administrator/perl5/lib/perl5';

print "Executing Report, wait please and not cancel this process...\n\n\n";
use Data::Dumper;
use DBI;
use Excel::Writer::XLSX;
use Email::Send::SMTP::Gmail;
use Number::Format;
use POSIX qw(strftime);
use Try::Tiny;
my $start_run = time();


# Connection config sybase
my $database = 'ACADSYS_USM';
my $server = '172.10.17.35';
my $port = 5001;
my $user = 'char1';
my $passwd = '654321';
my $dsn = "dbi:ODBC:DRIVER={FreeTDS};Server=$server;Port=$port;TDS_Version=5.0;database=$databyyase;";


my $to = 'hennyss1992@gmail.com';
my $cc = '';
my $bcc = '';
my $from = 'noresponder@usm.edu.ve';
my $subject2 = "Error al ejecutar script automatico";
my $epw = '/*sys20!7+-';#
my $mail=Email::Send::SMTP::Gmail->new( -smtp=>'smtp.gmail.com', -login=>$from, -pass=>$epw, -port=>465, -layer=>'ssl');#




my $dbh = DBI->connect("$dsn", "$user", "$passwd") or die "Unable for connect to server $DBI::errstr" unless $dbh;

my $sb_query1 = "DROP TABLE TMP_CPREINS";
$sth = $dbh->prepare($sb_query1);
$sth->execute() or die SendMail('Ocurrio un error');

print "******TMP_CPREINS has been drop****** \n";


my $sb_query = "CREATE TABLE TMP_CPREINS
(
   NUCLEO      char(2)      NOT NULL,
   T_CONCEPTO  int NOT NULL,
   RECIBO      numeric(8,0) NOT NULL,
   COD_CARRERA char(7)      NULL,
   CEDULA      varchar(20)  NOT NULL,
   PERIODO     char(6)      NULL,
   TIPO        char(1)      NULL,
   FECHA_TRANS datetime     NULL)
INSERT INTO TMP_CPREINS (NUCLEO, RECIBO, T_CONCEPTO, CEDULA, COD_CARRERA, PERIODO, TIPO, FECHA_TRANS)
SELECT A.NUCLEO, A.RECIBO, A.T_CONCEPTO, B.CEDULA, B.COD_CARRERA, '201702','X',B.FECHA_REGISTRO
FROM SFA_PAGO_RECIBO A, SFA_RECIBO_TAPA B
WHERE A.RECIBO=B.RECIBO AND ESTATUS_RECIBO='R' AND A.T_CONCEPTO IN (SELECT B.T_CONCEPTO FROM SCA_CONCEPTO B WHERE  B.TIPO='PRE_INS' AND T_STATUSB=1)
SELECT * FROM TMP_CPREINS ORDER BY CEDULA
SELECT B.RECIBO, A.*
FROM SIN_PREINSCRIPCION A, TMP_CPREINS B
WHERE (A.RECIBO = NULL OR A.RECIBO<>B.RECIBO)AND A.CEDULA = B.CEDULA AND A.COD_CARRERA1=B.COD_CARRERA AND A.ADMITIDO=0 AND A.T_STATUSB=1 AND A.PERIODO=B.PERIODO 
UPDATE SIN_PREINSCRIPCION SET RECIBO = B.RECIBO, FECHA_TRANS=B.FECHA_TRANS
FROM SIN_PREINSCRIPCION A, TMP_CPREINS B
WHERE (A.RECIBO = NULL OR A.RECIBO<>B.RECIBO)AND A.CEDULA = B.CEDULA AND A.COD_CARRERA1=B.COD_CARRERA AND A.ADMITIDO=0 AND A.T_STATUSB=1 AND A.PERIODO=B.PERIODO";



$sth = $dbh->prepare($sb_query);
$sth->execute() or die SendMail('Ocurrio un error');

print "                       PLEASE WAIT!!!!!!             \n\n\n";
print "*****TMP_CPREINS has been creater and writer******* \n";




sub SendMail { #this send correct mail info
 my @params = @_;
 $message = $params[0];
 print("ADVERTENCIA: ".$message." \n");#
 $mail->send(-to=>$to,
   -bcc=>$bcc,
   -subject=>$subject2,
   -verbose=>'1',
   -body=>"Ocurrio un error al generar el script",
   -contenttype=>'text/html'
   );
 $mail->bye;
}


my $end_run = time();
my $run_time = ($end_run - $start_run);
print "End Script: $run_time seconds\n";
