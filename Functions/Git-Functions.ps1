Function Get-GitStatus {
    & git status -sb $args
}
New-Alias -Name s -Value Get-GitStatus -Force -Option AllScope
Function Get-GitCommit {
    & git commit -ev $args
}
New-Alias -Name c -Value Get-GitCommit -Force -Option AllScope
Function Get-GitAdd {
    & git add --all $args
}
New-Alias -Name ga -Value Get-GitAdd -Force -Option AllScope
Function Get-GitTree {
    & git log --graph --oneline --decorate $args
}
New-Alias -Name t -Value Get-GitTree -Force -Option AllScope
Function Get-GitPush {
    & git push $args
}
New-Alias -Name gps -Value Get-GitPush -Force -Option AllScope
Function Get-GitPull {
    & git pull $args
}
New-Alias -Name gpl -Value Get-GitPull -Force -Option AllScope
Function Get-GitFetch {
    & git fetch $args
}
New-Alias -Name f -Value Get-GitFetch -Force -Option AllScope
Function Get-GitCheckout {
    & git checkout $args
}
New-Alias -Name co -Value Get-GitCheckout -Force -Option AllScope
Function Get-GitBranch {
    & git branch $args
}
New-Alias -Name b -Value Get-GitBranch -Force -Option AllScope
Function Get-GitRemote {
    & git remote -v $args
}
New-Alias -Name r -Value Get-GitRemote -Force -Option AllScope