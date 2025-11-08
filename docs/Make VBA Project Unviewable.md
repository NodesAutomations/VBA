Reference : https://www.linkedin.com/pulse/locking-excel-vba-projects-from-view-riley-carney/

### Overview
One of the most difficult parts of creating an Excel VBA project is protecting the code behind the project itself. While Microsoft includes a password-protection system that is relatively secure (though it can be broken hash-wise with about 2^11*50 iterations), it doesn't offer a reasonable solution to those who want to remove the password and understand the power of a quick Google search. Luckily, there is a more robust solution available to those who want to keep their project secure through editing the hex-values of the macro document.

### Open Excel File Contents
First off, we will need to download a hex-editor software off the internet. Many will work, but personally I prefer HxD, where you can find a link here:
Next up, you will want to change the extension of the Excel Macro document into a .ZIP in order to look into the back-end of an Excel document. For some computers, you may require an outside program in order to do this, and personally I would recommend WinRAR or 7Zip.

Next, you will want to navigate to the \xl\ folder of the file. In this, there will be a file named "vbaProject.bin." Unzip this file into another folder where you are able to manipulate it, and lets work the magic!
In your hex-editor, search for the value "CMG." This grouping of hex values is what we will change to make the project unable to be viewed for normal and intermediate computer users. Once you find it, it should look similar to this:

![image](https://user-images.githubusercontent.com/60865708/194349674-fe31b421-6c36-4f97-8b1e-c374fdcc2e2d.png)

### Modification to VBA File
> We will want to change everything in quotes for the variables "CMG", "DPB", as well as "GC", into "F". The amount that is changed needs to be either the same amount of characters as what's inside the quotations, or if you are modifying an excel macro file, it can also be greater than the characters. Just be sure not to go into the next variable and leave room for a quote at the end! After modifying the file we will get something that looks like:

![image](https://user-images.githubusercontent.com/60865708/194349742-1e228547-39be-4bc2-8e06-6acf6c099ba8.png)


Also, make sure to make a backup of the excel file before moving the .bin back into the Excel Macro file or else you will not be able to access your project. Have one file for development, and another file for production.

### Results
After, simply replace the old "vbaProject.bin" inside the \xl\ folder and change the Excel Macro file back into a ".xlsm" extension. Once we try to view the project inside the developer tab we are unable to view it, as it will give the error:

![image](https://user-images.githubusercontent.com/60865708/194349785-1917099f-b43b-4590-ac6b-e79aa1ef4d14.png)

While it isn't foolproof, this is a way to prevent those who are amateur and intermediate users access to the code inside your Excel file, and will allow you to know that your clients will not be able to easily change your code or view it under normal circumstances. It will allow you to breath easy knowing that they will not modify any code that you did not originally want modified unless they're an adept computer user.

 
# Excel Lock vba project from viewing****
### Goals

- Make VBA Code Unreadable for End Users

### Steps

- Convert Excel file to zip
- Open zip and find zip/xl/vbaProject.bin file
- Open That bin file
- Find CMG and DPG and GC Values and Replace with 1 + Add Additional 1
- New Key Length= Old Key Length +1
- Save and convert back to original excel file

### Before

```jsx
CMG="64666CA8F4P"
DPB="2A2822F6E9F7E9F7E9"
GC="F0F2F8FBF9FBF904"
```

### After

```jsx
CMG="11111111111"
DPB="1111111111111111111"
GC="11111111111111111"
```
