# Optimus
Loading automation tool that minimizes wasted space while maximizing quality of loads composition by pairing wisely inventory available and dealers' needs.

## General informations on algorithm procedure
Optimus algorithm proceeds executing the following steps:

- #### Search of **perfect match** between orders (**wishes**) from orders' list (**wishlist**) and available inventory
    
    We run through each **wish** beginning by the ones with lowest rank and associate it inventory if there's enough available to fulfill it. 
    
- #### Construction of loads from the association made in the **perfect match**

    For each plant to plant, we try to build the largest number of loads possible with items at our reach (coming from the **perfect match**). Loads built we'll be only considered if they respect all **constraints** defined by BRP. **Constraints are detailed in a further section**
    
    If the number of loads built exceeds the maximum accepted for the plant to plant, only the best ones in terms of highest number of **mandatory** crates and lowest **average ranking** will be kept
    
    Items that could not made it to the loads are sent back to the available inventory for the next step
    
- #### Satisfaction of minimums

  Using the inventory remaining from last step and considering **wishes** that are still unfulfilled, we run trough each plant to plant begining by the ones with the lowest priority number and try to build enough loads to satisfy their minimums expected if it wasn't reached earlier
   
  Items that could not made it to the loads are sent back to the available inventory for the next step
  
- #### Satisfaction of maximums

  Using the inventory remaining from last step and considering **wishes** that are still unfulfilled, we run trough each plant to plant begining by the ones with the lowest priority number and try to build enough loads to satisfy their maximums expected if it wasn't reached earlier
  
  Items that could not made it to the loads are sent back to the available inventory for the next step
  
- #### Distribution of leftovers

  Using the inventory remaining from the whole process and considering **wishes** that are still unfulfilled, we try to fill space remaining in truck trailers that are already built and ready to go
  
## Usage
Program can be run by clicking on `run_optimus.py` from the following path : S:\Shared\Business_Planning\Tool\Plant_to_plant\Optimus

