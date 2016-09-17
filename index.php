<?
echo '<html> <head>
  <meta charset="utf-8">
  <title>расписание "Норма"</title>
 </head>';
require_once '../wp-config.php'; // DB_NAME DB_PASSWORD DB_USER DB_HOST//
//избранные преподы для расшифровки
$prp=array(0=>"","КОВАЛЕНКО", "КОЖУХАР","МІЗЮК","ДУЩЕНКО","ДМИТРІЄВА");
//block date
$currenttime = time();
	$date_time_array = getdate($currenttime);
    $hours = $date_time_array['hours'];
    $minutes = $date_time_array['minutes'];
    $seconds = $date_time_array['seconds'];
    $month = $date_time_array['mon'];
    $day = $date_time_array['mday'];
    $year = $date_time_array['year'];
	$day+=7;
	$timestamp= mktime($hours,$minutes,$seconds,$month,$day,$year);
	$date2=strftime('%Y-%m-%d',$timestamp);
	$week=array(0=>"воскресенье", "понедельник","вторник","среда","четверг","пятница","суббота");
//   
$host=DB_HOST; // имя хоста (уточняется у провайдера)
$database=DB_NAME; // имя базы данных, которую вы должны создать
$user=DB_USER; // заданное вами имя пользователя, либо определенное провайдером
$pswd=DB_PASSWORD; // заданный вами пароль
 $dbh = mysql_connect($host, $user, $pswd) or die("Не могу соединиться с MySQL.");
mysql_select_db($database) or die("Не могу подключиться к базе.");
// only groups
$query = "SELECT distinct gr FROM 1rozklad";
$res = mysql_query($query);
while($row = mysql_fetch_array($res))
{echo '&nbsp;<a href="?gr='.$row[0].'">-'.$row[0]."-</a>&nbsp;";}
echo ";)"; 
if($_GET["gr"] ||$_GET["prp"])
{
	$gr=(int)$_GET["gr"];
	$prpd=(int)$_GET["prp"];
if($gr>0) $usl="gr=$gr";
if($prpd>0) {$usl="prpd like '%$prp[$prpd]%'";}
	//echo $usl."ok!";
	//echo $date2."-time";
$query = "	SELECT displn,prpd, dat, nomer, gr FROM 1rozklad where $usl and dat between STR_TO_DATE(now(),'%Y-%m-%d') and STR_TO_DATE('".$date2."','%Y-%m-%d') order by dat,nomer";
//echo $query;
$res = mysql_query($query);
$tab='<table>';
while($row = mysql_fetch_array($res))
{
	$dt=explode("-",$row[2]);
	$nom_day=date("w",mktime(0,0,0,$dt[1],$dt[2],$dt[0]));
	if ($oldnom_day!=$nom_day)	$day='bgcolor="#CCFFCC">'.$week[$nom_day]."<BR>".$row[2]; else $day=">";
	//$day=$week[$nom_day];
	$displ=trim($row[0]);
	if(mb_strlen($displ)>6) $displ=mb_substr($displ, 0, 1).mb_substr(mb_strtolower($displ),1);
	if($row[3]==1) $npara='bgcolor="#FFAAAA"';else $npara='';
$tab.='<tr><td '.$day.'</td><td '.$npara.'>'.$row[3].' пара</td><td bgcolor="#AAAAFF">'.$row[4].'</td><td>'.$displ.'</td><td>'.$row[1].'</td></tr>';	
$oldnom_day=$nom_day;
//Я ТАК думаю что зоздаем массив с ключами "день недели-номер пары" и запихиваем в шаблон
}
echo $tab.'</table>';
}
for($i=1;$i<count($prp);$i++)
{echo '&nbsp;<a href="?prp='.$i.'">-'.$prp[$i]."-</a>&nbsp;";}
echo '</html>';
?>