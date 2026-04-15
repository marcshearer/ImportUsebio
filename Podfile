# Uncomment the next line to define a global platform for your project
platform :macos, '10.13'

target 'ImportUsebio' do
  # Comment the next line if you don't want to use dynamic frameworks
  use_frameworks!

  # Pods for ImportUsebio
  pod 'libxlsxwriter', '~> 1.1'

end

# Make minimum target at least 10.13 to avoid arclite error
post_install do |installer|
  installer.generated_projects.each do |project|
    project.targets.each do |target|
      target.build_configurations.each do |config|
        # Sets the minimum deployment target to 10.13 for all pod targets
        config.build_settings['MACOSX_DEPLOYMENT_TARGET'] = '10.13'
      end
    end
  end
end
