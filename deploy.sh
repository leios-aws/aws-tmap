#!/bin/sh
NAME=$(basename $(realpath .))
REVISION=$(git describe --all | cut -d '/' -f 2)

rm -f ${NAME}.zip

case ${REVISION} in
    *develop*)
        cd nodejs && \
            zip -r ../${NAME}.zip . -x node_modules\* && \
            cd .. && \
            aws lambda update-function-code --function-name ${NAME} --zip-file fileb://${NAME}.zip
        ;;
    *)
        cd nodejs && \
            zip -r ../${NAME}.zip . -x node_modules\* && \
            cd .. && \
            aws lambda update-function-code --function-name ${NAME} --zip-file fileb://${NAME}.zip
	    for RETRY in 1 2 3 4 5
        do
            sleep 2
            RESULT=$(aws lambda publish-version --function-name ${NAME})
            STATUS=$(echo ${RESULT} | jq -r .LastUpdateStatus)
            if [ "${STATUS}" = "Successful" ]
            then
                echo "STATUS: ${STATUS}"
                VERSION=$(echo ${RESULT} | jq -r .Version)
                break
            fi
        done
        aws lambda update-alias --function-name ${NAME} --function-version ${VERSION} --name service
        ;;
esac
