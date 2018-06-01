<!DOCTYPE html>
<html lang="en">
<head>
    <title>Excel Generator</title>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css">
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script>
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>
</head>
<body>
<div id="main-div">
    <form class="form-inline" action="ExcelGenerator.php" method="post">
        <div class="form-group">
            <label for="coin">Coin :</label>
            <input type="text" class="form-control" id="coin" name="coinSymbolFrom">
        </div>
        <div class="form-group">
            <select name="coinSymbolTo" class="form-control">
                <option value="BTC">BTC</option>
                <option value="ETH">ETH</option>
                <option value="XRM">XRM</option>
                <option value="USDT">USDT</option>
            </select>
        </div>
        <button type="submit" class="btn btn-primary">Submit</button>
    </form>
</div>
</body>
<style>
    #main-div{
        height:auto;
        margin:100px auto auto auto;
    }

    form{
        width:50%;
        margin:0 auto;
    }
</style>
</html>