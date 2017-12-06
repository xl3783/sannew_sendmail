<?php

require_once("PHPMailer/class.phpmailer.php");
require_once("PHPMailer/class.smtp.php");
require_once("PHPExcel/PHPExcel.php");
require_once("smarty/Smarty.class.php");

$file = dirname(__FILE__).'\source\config.xlsx';
$objPHPExcel = PHPExcel_IOFactory::load($file);
$sheetData = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);

$mailUrl = $sheetData[2]['A'];//邮箱
$mailPwd = $sheetData[2]['C'];//密码
$mailUname = $sheetData[2]['B'];//昵称
$time = $sheetData[2]['D'];
$time2 = $sheetData[2]['D'];
$time = ($time - 25569)*24*60*60;
$year = gmdate('Y年',$time);
$month = gmdate('m月',$time);
$day = gmdate('d日',$time);

$dataFile = $sheetData[2]['G'];
$title = $sheetData[2]['E'];//标题
$theme = $sheetData[2]['F'];//主题
$file = dirname(__FILE__)."\\source\\".$dataFile;
$file = iconv("utf-8","gb2312",$file);
$objPHPExcel = PHPExcel_IOFactory::load($file);
$sheetData = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);

$smarty = new Smarty();
$number = $name=$department = $duty = $baseSalary=$performanceSalary=$storAllowance=$fullDuty=$hiringAward=$classFee=$otherAllowance=$houseAllowance
=$monthWage=$leaveDays=$leaveCost=$sickHours=$sickCost=$otherCost=$totalCost=$factWage=$socialSecurity=$houseFund=$tax=$totalWage=$higherupGrade=
    $innerGrade=$outerGrade=$finnalGrade='';
array_shift($sheetData);
array_shift($sheetData);

foreach ($sheetData as $row){
    if(null == $row['A']){
        break;
    }
    $number = $row['A'];//无工号对应项
    $name = $row['B'];
    $department = $row['C'];
    $duty = $row['D'];
    $baseSalary = $row['E'];
    $performanceSalary = $row['F'];
    $storAllowance = $row['G'];
    $fullDuty = $row['J'];
    $hiringAward = $row['H'];
    $classFee = $row['I'];
    $otherAllowance = $row['K'];
    $houseAllowance = $row['L'];
    $monthWage = $row['M'];
    $leaveDays = $row['N'];
    $leaveCost = $row['O'];
    $sickHours = $row['P'];
    $sickCost = $row['Q'];
    $otherCost = $row['R'];
    $totalCost = $row['S'];
    $factWage = $row['T'];
    $socialSecurity = $row['U'];
    $houseFund = $row['V'];
    $tax = $row['X'];
    $totalWage = $row['Y'];
    $higherupGrade = $row['AC'];
    $innerGrade = $row['AD'];
    $outerGrade = $row['AE'];
    $finnalGrade = $row['AF'];

    $data = [
        'number'=>$number,
        'name'=>$name,
        'department'=>$department,
        'duty'=>$duty,
        'baseSalary'=>$baseSalary,
        'performanceSalary'=>$performanceSalary,
        'storAllowance'=>$storAllowance,
        'fullDuty'=>$fullDuty,
        'hiringAward'=>$hiringAward,
        'classFee'=>$classFee,
        'otherAllowance'=>$otherAllowance,
        'houseAllowance'=>$houseAllowance,
        'monthWage'=>$monthWage,
        'leaveDays'=>$leaveDays,
        'leaveCost'=>$leaveCost,
        'sickHours'=>$sickHours,
        'sickCost'=>$sickCost,
        'otherCost'=>$otherCost,
		'totalCost'=>$totalCost,
        'factWage'=>$factWage,
        'socialSecurity'=>$socialSecurity,
        'houseFund'=>$houseFund,
        'tax'=>$tax,
        'totalWage'=>$totalWage,
        'higherupGrade'=>$higherupGrade,
        'innerGrade'=>$innerGrade,
        'outerGrade'=>$outerGrade,
        'finnalGrade'=>$finnalGrade,
        'time'=>$time,
        'year'=>$year,
        'month'=>$month,
        'day'=>$day,
        'title'=>$title
    ];
    //$content = $smarty->fetch('工资条模板.html');
    $smarty->assign($data);
    $content = $smarty->fetch('template.html');
	
    if(sendMail($row['AB'],$theme,$content)){
        echo $name.':邮件发送成功!<br>';
    }else{
        echo $name.':邮件发送失败!<br>';
    }
}


/*发送邮件方法
 *@param $to：接收者 $theme：标题 $content：邮件内容
 *@return bool true:发送成功 false:发送失败
 */
function sendMail($to,$theme,$content){
    global $mailUrl, $mailPwd,$mailUname;
    //实例化PHPMailer核心类
    $mail = new PHPMailer();

    //是否启用smtp的debug进行调试 开发环境建议开启 生产环境注释掉即可 默认关闭debug调试模式
    $mail->SMTPDebug = 0;

    //使用smtp鉴权方式发送邮件
    $mail->isSMTP();

    //smtp需要鉴权 这个必须是true
    $mail->SMTPAuth=true;

    //链接qq域名邮箱的服务器地址
    $mail->Host = 'smtp.exmail.qq.com';

    //设置使用ssl加密方式登录鉴权
    $mail->SMTPSecure = 'ssl';

    //设置ssl连接smtp服务器的远程服务器端口号，以前的默认是25，但是现在新的好像已经不可用了 可选465或587
    $mail->Port = 465;

    //设置smtp的helo消息头 这个可有可无 内容任意
    // $mail->Helo = 'Hello smtp.qq.com Server';

    //设置发件人的主机域 可有可无 默认为localhost 内容任意，建议使用你的域名
    $mail->Hostname = 'localhost';

    //设置发送的邮件的编码 可选GB2312 我喜欢utf-8 据说utf8在某些客户端收信下会乱码
    $mail->CharSet = 'UTF-8';

    //设置发件人姓名（昵称） 任意内容，显示在收件人邮件的发件人邮箱地址前的发件人姓名
    $mail->FromName = $mailUname;

    //smtp登录的账号 这里填入字符串格式的qq号即可
    $mail->Username =$mailUrl;

    //smtp登录的密码 使用生成的授权码（就刚才叫你保存的最新的授权码）
    $mail->Password = $mailPwd;

    //设置发件人邮箱地址 这里填入上述提到的“发件人邮箱”
    $mail->From = $mailUrl;

    //邮件正文是否为html编码 注意此处是一个方法 不再是属性 true或false
    $mail->isHTML(true);

    //设置收件人邮箱地址 该方法有两个参数 第一个参数为收件人邮箱地址 第二参数为给该地址设置的昵称 不同的邮箱系统会自动进行处理变动 这里第二个参数的意义不大
    $mail->addAddress($to,' ');

    //添加多个收件人 则多次调用方法即可
     //$mail->addAddress('llxiang@sannewschool.com',' ');

    //添加该邮件的主题
    $mail->Subject = $theme;

    //添加邮件正文 上方将isHTML设置成了true，则可以是完整的html字符串 如：使用file_get_contents函数读取本地的html文件
    $mail->Body = $content;

    //为该邮件添加附件 该方法也有两个参数 第一个参数为附件存放的目录（相对目录、或绝对目录均可） 第二参数为在邮件附件中该附件的名称
    // $mail->addAttachment('./d.jpg','mm.jpg');
    //同样该方法可以多次调用 上传多个附件
    // $mail->addAttachment('./Jlib-1.1.0.js','Jlib.js');

    $status = $mail->send();

    //简单的判断与提示信息
    if($status) {
        return true;
    }else{
        return false;
    }
}

function read($filename){
    $objPHPExcel = PHPExcel_IOFactory::load($filename);
    $sheetData = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);
    return $sheetData;
}