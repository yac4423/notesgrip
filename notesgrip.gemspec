# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'notesgrip/version'

Gem::Specification.new do |spec|
  spec.name          = "notesgrip"
  spec.version       = Notesgrip::VERSION
  spec.authors       = ["yac4423"]
  spec.email         = ["notesgrip@tech-notes.dyndns.org"]
  spec.summary       = %q{Control Notes/Domino from Ruby Script.}
  spec.description   = %q{Control Notes/Domino from Ruby Script.}
  spec.homepage      = "https://github.com/yac4423/notesgrip"
  spec.license       = "MIT"

  spec.files         = `git ls-files -z`.split("\x0")
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.add_development_dependency "bundler", "~> 1.6"
  spec.add_development_dependency "rake", "~> 10.0"
end
