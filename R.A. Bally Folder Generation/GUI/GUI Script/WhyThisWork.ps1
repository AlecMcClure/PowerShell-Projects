$user_path = read-host -Prompt "Enter A File Path"
$id = $user_path
[regex]::Escape(("$id"))