Global setup:
 Download and install Git
  git config --global user.name "yuguorong"
  git config --global user.email yuguorong@gmail.com
        Next steps:
  mkdir xhzw
  cd xhzw
  git init
  touch README
  git add README
  git commit -m 'first commit'
  git remote add origin git@github.com:YuGuorong/xhzw.git
  git push origin master
      Existing Git Repo?
  cd existing_git_repo
  git remote add origin git@github.com:YuGuorong/xhzw.git
  git push origin master
      Importing a Subversion 