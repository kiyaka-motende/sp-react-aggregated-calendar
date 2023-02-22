echo
echo -e "\e[1;94mInstalling Node dependencies\e[0m"
CI_SERVER_HOST=npm.pkg.github.com

echo "registry=https://registry.npmjs.org/" > .npmrc
echo "@kiyaka-motende:registry=https://${CI_SERVER_HOST}" >> .npmrc
echo "//${CI_SERVER_HOST}/:_authToken=${CI_JOB_TOKEN}" >> .npmrc
echo "always-auth=true" >> .npmrc

npm install --save @kiyaka-motende/spfx-controls-react-sp-events@1.0.1
npm install

## commands to create dev certificate and copy it to the root folder of the project

echo
echo -e "\e[1;94mGenerating dev certificate\e[0m"
gulp trust-dev-cert

# Convert the generated PEM certificate to a CER certificate

openssl x509 -inform PEM -in ~/.rushstack/rushstack-serve.pem -outform DER -out ./spfx-dev-cert.cer

# Copy the PEM ecrtificate for non-Windows hosts

cp ~/.rushstack/rushstack-serve.pem ./spfx-dev-cert.pem

## add \*.cer to .gitignore to prevent certificates from being saved in repo

# if ! grep -Fxq '_.cer' ./.gitignore
# then
# echo "# .CER Certificates" >> .gitignore
# echo "_.cer" >> .gitignore
# fi

# ## add \*.pem to .gitignore to prevent certificates from being saved in repo

# if ! grep -Fxq '_.pem' ./.gitignore
# then
# echo "# .PEM Certificates" >> .gitignore
# echo "_.pem" >> .gitignore
# fi

echo
echo -e "\e[1;92mReady!\e[0m"

echo -e "\n\e[1;94m****\*\*****\nOptional: if you plan on using gulp serve, don't forget to add the container certificate to your local machine. Please visit https://aka.ms/spfx-devcontainer for more information\n****\*\*****"
