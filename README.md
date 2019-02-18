# SPUtility-JS


## How to Use this library
[SPUtility JS](https://github.com/ajDon/SPUtility-JS/blob/master/dist/SPUtility.min.js) add reference to the page also add JQuery to the page.

        <script  src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script>
        <script  src="SPUtility.min.js"></script>


## How to Use this library

1. Set Type Of Header

    As we have different header in OData like ```verbose```, ```minimalmetadata``` or ```nometadata```. Default header is verbose but you can change it.

        // To Set minimalmetadata
        SPUtility.setMinimalMetaData(true)

        // To Set nometadata
        SPUtility.setNoMetaData(true)

        // To Set to default change it to false
        SPUtility.setNoMetaData(false)
            OR
        SPUtility.setMinimalMetaData(false)

2. Set Site URL

        SPUtility.setSiteUrl("https://<siteName>.sharepoint.com")

3. Get User Properties

        SPUtility.getUserProperties('testUser@domain.com',"PropertyName").then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

4. Get All User Properties

        SPUtility.getAllUserProperties('testUser@domain.com').then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

5. Current User

        SPUtility.web.currentUser().then(function(user){
            console.info(user);
        },function(error){
            console.error(error);
        })

6. Ensure User

        SPUtility.web.ensureUser('testUser@domain.com').then(function(user){
            console.info(user);
        },function(error){
            console.error(error);
        })

7. Get Roles

        SPUtility.web.getRoles().then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

8. Check web permission inheritance is broken or not

        SPUtility.web.hasuniqueroleassignments().then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

9. Break Permission for Web

        var copyRoleAssignment = true, clearSubscope = false;
        SPUtility.web.breakroleinheritance(copyRoleAssignment,clearSubscope).then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

10. Reset Permission for Web

        SPUtility.web.resetroleinheritance().then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

11. Add Permission for Web

        var groupIdorUserId = 1, roleDefId = 1;
        SPUtility.web.addRoleAssignment(groupIdorUserId,roleDefId).then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

12. Remove Permssion from Web

        var groupIdorUserId = 1;
        SPUtility.web.deleteRoleAssignment(groupIdorUserId).then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

13. Check List has unique permission

        SPUtility.web.lists.getListByTitle("ListTitle").hasuniqueroleassignments().then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

        OR

        SPUtility.web.lists.getListById("60a3a6aa-1947-4806-85e5-5d6718f194d3").hasuniqueroleassignments().then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

14. Break Permission for list

        var copyRoleAssignment = true, clearSubscope = false;

        SPUtility.web.lists.getListByTitle("ListTitle").breakroleinheritance(copyRoleAssignment,clearSubscope).then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

        OR

        SPUtility.web.lists.getListById("60a3a6aa-1947-4806-85e5-5d6718f194d3").breakroleinheritance(copyRoleAssignment,clearSubscope).then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })