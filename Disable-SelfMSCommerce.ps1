Import-Module -Name MSCommerce
Connect-MSCommerce

$products = Get-MSCommerceProductPolicies -PolicyId AllowSelfServicePurchase

foreach ($product in $products) {
    Update-MSCommerceProductPolicy -PolicyId AllowSelfServicePurchase -ProductId $product.ProductID -Enabled $False   
}