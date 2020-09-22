<div align="center">

## Quick Tab Order


</div>

### Description

Quick tip for setting control tab order quickly.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dale E\. Lucid](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dale-e-lucid.md)
**Level**          |Beginner
**User Rating**    |3.6 (18 globes from 5 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VBA MS Access
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dale-e-lucid-quick-tab-order__1-23587/archive/master.zip)





### Source Code

```
A speedy way to set the tab order is to work thru the controls in reverse order setting the tab stop property to 0 on each. When you reach the first control the tab order is set.
When setting the same property on successive controls there is no need to bounce from the form to the property page. Once the focus is set to that property the first time it remains on that property (if available) with each object that receives focus from that point. In the above example setting the tab stop would be... control.. 0, control... 0, control... 0.
To get the effect of setting focus to a control by clicking on it's label, as in Access, simply set the tab stop for the label as the next in order before the control that it labels. Since a label can't receive focus at run time the focus goes to the next object in tab order that can receive focus.
```

