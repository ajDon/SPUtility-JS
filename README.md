# SPUtility-JS

## How to Use this library

Add [SPUtility JS](https://github.com/ajDon/SPUtility-JS/blob/master/dist/SPUtility.min.js) reference to the page.

        <script  src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script>
        <script  src="SPUtility.min.js"></script>

## How to Use this library

1.  Set Type Of Header

    As we have different header in OData like `verbose`, `minimalmetadata` or `nometadata`. Default header is verbose but you can change it.

        // To Set minimalmetadata
        SPUtility.setMinimalMetaData(true)

        // To Set nometadata
        SPUtility.setNoMetaData(true)

        // To Set to default change it to false
        SPUtility.setNoMetaData(false)
            OR
        SPUtility.setMinimalMetaData(false)

2.  Set Site URL

        SPUtility.setSiteUrl("https://<siteName>.sharepoint.com")

3.  Get User Properties

        SPUtility.getUserProperties('testUser@domain.com',"PropertyName").then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

4.  Get All User Properties

        SPUtility.getAllUserProperties('testUser@domain.com').then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

5.  Current User

        SPUtility.web.currentUser().then(function(user){
            console.info(user);
        },function(error){
            console.error(error);
        })

6.  Ensure User

        SPUtility.web.ensureUser('testUser@domain.com').then(function(user){
            console.info(user);
        },function(error){
            console.error(error);
        })

7.  Get Roles

        SPUtility.web.getRoles().then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

8.  Check web permission inheritance is broken or not

        SPUtility.web.hasuniqueroleassignments().then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

9.  Break Permission for Web

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

15. Add Permission to list

        var groupIdorUserId = 1, roleDefId = 1;

        SPUtility.web.lists.getListByTitle("ListTitle").addRoleAssignment(groupIdorUserId,roleDefId).then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

        OR

        SPUtility.web.lists.getListById("60a3a6aa-1947-4806-85e5-5d6718f194d3").addRoleAssignment(groupIdorUserId,roleDefId).then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

16. Reset Permission for list

        SPUtility.web.lists.getListByTitle("ListTitle").resetroleinheritance().then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

        OR

        SPUtility.web.lists.getListById("60a3a6aa-1947-4806-85e5-5d6718f194d3").resetroleinheritance().then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

17. Remove Permission from list

        var groupIdorUserId = 1;
        SPUtility.web.lists.getListByTitle("ListTitle").deleteRoleAssignment(groupIdorUserId).then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

        OR

        SPUtility.web.lists.getListById("60a3a6aa-1947-4806-85e5-5d6718f194d3").deleteRoleAssignment(groupIdorUserId).then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

18. Get Default Form URL for List

        // Get Default display Form Url, similarly with get by list id.
        SPUtility.web.lists.getListByTitle("ListTitle").getDefaultDisplayFormUrl().then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

        // Get Default edit Form Url, similarly with get by list id.
        SPUtility.web.lists.getListByTitle("ListTitle").getDefaultEditFormUrl().then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

        // Get Default new Form Url, similarly with get by list id.
        SPUtility.web.lists.getListByTitle("ListTitle").getDefaultNewFormUrl().then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

19. Get Items from list

        //Similarly with get by list id.
        SPUtility.web.lists.getListByTitle("ListTitle").Items
                .select(["Title","Autor/Title","Autor/Id"])
                .expand(["Author"]).topItem(100)
                .filter("Title eq 'Test'")
                .get().then(function(data){
                    console.info(data);
                },function(error){
                    console.error(error);
                })

20. Get All items from list

        //Similarly with get by list id.
        SPUtility.web.lists.getListByTitle("ListTitle").Items
                .select(["Title","Autor/Title","Autor/Id"])
                .expand(["Author"])
                .filter("Title eq 'Test'")
                .getAllItems().then(function(data){
                    console.info(data);
                },function(error){
                    console.error(error);
                })

21. Get Items as paged from list

        //Similarly with get by list id.
        SPUtility.web.lists.getListByTitle("ListTitle").Items
                .select(["Title","Autor/Title","Autor/Id"])
                .expand(["Author"]).topItem(100)
                .filter("Title eq 'Test'")
                .getPaged().then(function(data,nextItem){
                    console.info(data);
                    if(nextItem){
                        nextItem.getNextItems().then(function(data,nextItem){
                            console.info(data);
                        },function(error){
                            console.error(error);
                        })
                    }
                },function(error){
                    console.error(error);
                })

22. Get item by id from list

        //Similarly with get by list id.
        SPUtility.web.lists.getListByTitle("ListTitle").Items
                .select(["Title","Autor/Title","Autor/Id"])
                .expand(["Author"])
                .getItemById(1).then(function(data){
                    console.info(data);
                },function(error){
                    console.error(error);
                })

23. Add new item to list

        //Similarly with get by list id.
        SPUtility.web.lists.getListByTitle("ListTitle").Items
                .addNewItem({Title: 'Test',LookupId: 23}).then(function(data){
                    console.info(data);
                },function(error){
                    console.error(error);
                })

24. Update item in list

        var itemId = 1, eTag = "'1'";
        //Similarly with get by list id.
        SPUtility.web.lists.getListByTitle("ListTitle").Items
                .updateItem(itemId, eTag, {Title: 'Test',LookupId: 23}).then(function(data){
                    console.info(data);
                },function(error){
                    console.error(error);
                })

25. Delete item from list

        var itemId = 1, eTag = "'1'";
        //Similarly with get by list id.
        SPUtility.web.lists.getListByTitle("ListTitle").Items
                .deleteItem(itemId, eTag).then(function(data){
                    console.info(data);
                },function(error){
                    console.error(error);
                })

26. Check item level permission is breaked or not

        var itemId = 1;
        //Similarly with get by list id.
        SPUtility.web.lists.getListByTitle("ListTitle").Items
                .hasuniqueroleassignments(itemId).then(function(data){
                    console.info(data);
                },function(error){
                    console.error(error);
                })

27. Break Permission at item level

        var itemId = 1, copyRoleAssignment = true, clearSubscope = false;
        //Similarly with get by list id.
        SPUtility.web.lists.getListByTitle("ListTitle").Items
                .breakroleinheritance(itemId , copyRoleAssignment, clearSubscope)
                .then(function(data){
                    console.info(data);
                },function(error){
                    console.error(error);
                })

28. Reset Permission for item

        var itemId = 1;
        //Similarly with get by list id.
        SPUtility.web.lists.getListByTitle("ListTitle").Items
                .resetroleinheritance(itemId).then(function(data){
                    console.info(data);
                },function(error){
                    console.error(error);
                })

29. Add Permission to list item

        var itemId = 1, groupIdorUserId = 1, roleDefId = 1;
        //Similarly with get by list id.
        SPUtility.web.lists.getListByTitle("ListTitle").Items
            .addRoleAssignment(itemId, groupIdorUserId, roleDefId).then(function(data){
                console.info(data);
            },function(error){
                console.error(error);
            })

30. Remove user/group permission from list item

        var itemId = 1, groupIdorUserId = 1;
        //Similarly with get by list id.
        SPUtility.web.lists.getListByTitle("ListTitle").Items
            .deleteRoleAssignment(itemId, groupIdorUserId).then(function(data){
                console.info(data);
            },function(error){
                console.error(error);
            })

31. Get Fields from List

        //Similarly with get by list id.
        SPUtility.web.lists.getListByTitle("ListTitle").Fields
            .select(['Title','InternalName'])
            .filter('Hidden eq false')
            .get().then(function(data){
                console.info(data);
            },function(error){
                console.error(error);
            })

32. Get Site Columns from Web

        SPUtility.web.siteColumns().Fields
            .select(['Title','InternalName'])
            .filter('Hidden eq false')
            .get().then(function(data){
                console.info(data);
            },function(error){
                console.error(error);
            })

33. Batch execution for adding, updating or deleting item in list

        var batch = new SPUtility.createNewBatch();
        
        SPUtility.web.lists.getListByTitle('ListTitle').Items
            .addNewItemUsingBatch({Title: 'Test'},batch)
        
        var updateItemId = 2, itemETag = "*";
        SPUtility.web.lists.getListByTitle('ListTitle').Items
            .updateItemUsingBatch(updateItemId, itemETag, {Title: 'Test'},batch)

        SPUtility.web.lists.getListByTitle('ListTitle').Items
            .updateItemUsingBatch(40, "'1'", {Title: 'Test'},batch)

        var deleteItemId = 3, deleteItemETag = "*";
        SPUtility.web.lists.getListByTitle('ListTitle').Items
            .deleteItemUsingBatch(updateItemId, itemETag, {Title: 'Test'},batch)

        batch.executeBatch().then(function(data){
            console.info(data);
        },function(error){
            console.error(error);
        })

34. Site Groups

        SPUtility.web.siteGroups.groups
            .select(['Title','Id'])
            .filter("Id gt 20")
            .getAllGroups().then(function(data){
                    console.info(data);
                },function(error){
                    console.error(error);
                })
        
        
