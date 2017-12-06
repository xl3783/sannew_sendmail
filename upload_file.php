<?php
/**
 * Created by PhpStorm.
 * User: admin
 * Date: 2017/12/6
 * Time: 11:14
 */
//var_dump($_FILES);exit();
//
$allowedExts = array("xlsx");
$temp = explode(".", $_FILES["file"]["name"]);
$extension = end($temp);     // 获取文件后缀名
if ((($_FILES["file"]["type"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
    && ($_FILES["file"]["size"] < 1048576)   // 小于 2M
    && in_array($extension, $allowedExts))
{
    if ($_FILES["file"]["error"] > 0)
    {
        echo "错误：: " . $_FILES["file"]["error"] . "<br>";
    }
    else
    {
        move_uploaded_file($_FILES["file"]["tmp_name"], "source/userdata.xlsx");
    }
}
else
{
    echo "请上传.xlsx格式的文件！";
}