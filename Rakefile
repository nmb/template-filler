task :versionfile do
  version = `git log -1 --format="%h %ad"`
  File.open('version.rb', 'w') do |f|
    f.write("@version=\"Version: \n#{version}\"\n")
  end
end

task :build => :versionfile do
  sh 'ocra --dll ruby_builtin_dlls/libssp-0.dll --dll ruby_builtin_dlls/libgmp-10.dll --dll ruby_builtin_dlls/libffi-7.dll --dll ruby_builtin_dlls\zlib1.dll --dll ruby_builtin_dlls\libssl-1_1-x64.dll --dll ruby_builtin_dlls/libcrypto-1_1-x64.dll --gem-full=openssl --gem-all=fiddle --gem-all=libui --add-all-core --output template-filler.exe --gem-all --windows main.rb fill.rb version.rb'
end

