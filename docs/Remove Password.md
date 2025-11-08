Up until this research, the change to make a VBA Project locked/unviewable was said to be irreversible but I discovered that if you simulate a password protected document by setting the *[ProjectCLSID](https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/20f4aad3-b646-4311-8301-5948fb1c2ded)* to all zero’s and use valid values for *ProjectProtectionState* (CMG), *ProjectVisibilityState* (GC) and *ProjectPassword* (DPB) you can undo this protection.

**ID={00000000–0000–0000–0000–000000000000}**

**CMG=CAC866BE34C234C230C630C6DPB=94963888C84FE54FE5B01B50E59251526FE67A1CC76C84ED0DAD653FD058F324BFD9D38DED37GC=5E5CF2C27646414741474**

Above are values that will undo the protection, but because the MS Office Compound File Binary Format (CFBF) is sensitive to data length changes, your best bet is to let [EvilClippy](https://github.com/outflanknl/EvilClippy) make these changes for you.

Bonus: The EvilCippy ‘-uu’ option also removes any password protection from the VBA Project.

Ref : [VBA Project Locked; Project is Unviewable | by Carrie Roberts | Walmart Global Tech Blog | Medium](https://medium.com/walmartglobaltech/vba-project-locked-project-is-unviewable-4d6a0b2e7cac)
