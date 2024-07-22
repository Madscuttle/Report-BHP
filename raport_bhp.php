<html>
<center>
<meta charset="utf-8">
<div align="center">
<form action="../raporty/raport_bhp.php" method="post">

<label for="datestart">Od:</label>
<input type="date" name="datestart"  value="2020-01-01" required>
<label for="dateend">Do:</label>
<input type="date" name="dateend"  value="<?php echo date('Y-m-d'); ?>" required><br /><br />

<input type="submit" name="wyszukaj" value="EXCEL" />

</form>
</div>

<?php
	if(isset($_POST['wyszukaj'])) { 
		$param1=$_POST['datestart'];
		$param2=$_POST['dateend'];
		$output = exec("python raport_bhp.py '$param1' '$param2'");
		
		echo $output;
		
		header('Content-Description: File Transfer');
		header('Content-Type: application/octet-stream');
		header('Content-Disposition: attachment; filename='.basename($output));
		header('Content-Transfer-Encoding: binary');
		header('Expires: 0');
		header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
		header('Pragma: public');
		header('Content-Length: ' . filesize($output));
		ob_clean();
		flush();
		readfile($output);
		unlink($output);
		header("Location: index.php");
		exit();
		
	
	}	
	?>

</center>
</html>