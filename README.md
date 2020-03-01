# Optimus <img src=Readme_Pictures/bombardier-recreational-products-brp-vector-logo-small.png width="55">
Loading automation tool that minimizes wasted space while maximizing quality of loads composition by pairing wisely inventory available and dealers' needs.

## General informations on algorithm procedure for the main process
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
Program can be run by clicking on `run_optimus.py` at the following path : **S:\Shared\Business_Planning\Tool\Plant_to_plant\Optimus**

An interface will open and allow user to select among three different modes of operations

1. **P2P Full Process**
2. **Fast Loads**
3. **Forecast**

<p align="center">
<img src=Readme_Pictures/Mode.png>
</p>

### P2P Full process
This Optimus mode drives all daily morning operations linked to loads buidling.

Once chosen on the **Mode selection** interface, the full process starts by opening another interface leaving the user a lot of flexibility and options for the loads building of every plant to plant.

<p align="center">
<img src=Readme_Pictures/p2p_param.png>
</p>

#### **Max and min**
As seen above, we can set a minimum and a maximum number of loads that we want to respect for each plant to plant.
It's important to be aware of the following details:

- The minimum can be set to 0 without any problem. If it's greater than 0, Optimus will try to fulfill this constraint as much as it can

- If the user decides to do not build any loads for one plant to plant, he must press the **SKIP** button. The **SKIP** button will ensure that every orders (**wish**) linked to this plant to plant will not be considered in the process. 

- The maximum of two plant to plant are currently shared if both plant to plant have the same **POINT TO** and their **POINT FROM** are either **4100** or **4125**. This shared maximum number of loads will always be strictly respected.

#### **Flatbed**

The flatbed presented in the interface refer to the 48 feet long trailer. However, all plant to plant with **POINT FROM** that are either **4100** or **4125** share also implicitly to 53 feet long flatbed at anytime.

#### **Priority**

The priority indicates which plant to plant to consider first while satisfying maximums and minimums.
One priority number can be use more than once.

#### **Transit**

Transit is actually not considered by the algorithm






    

