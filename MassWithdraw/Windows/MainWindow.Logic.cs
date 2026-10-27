using System;
using System.Linq;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Lumina.Excel;
using Dalamud.Plugin.Services;
using GameStructs = FFXIVClientStructs.FFXIV.Client.Game;
using ItemRow = Lumina.Excel.Sheets.Item;
using Dalamud.Game;
using FFXIVClientStructs.FFXIV.Client.Game;

namespace MassWithdraw.Windows;

public partial class MainWindow
{

#region Config & IDs
    private const uint
        RareGearId           = 999001,
        WhiteGearId          = 999002,
        MateriaId            = 999003,
        ConsumablesId        = 999004,
        RegisterableId      = 999005,
        GearBoxId           = 999006,
        CatalystId          = 999007,
        CraftingMaterialsId  = 999008,
        CurrencyOtherId  = 999009,
        TreasureMaterialsId = 999010,
        SubmersiblePartsId   = 999011,
        SchematicBoardId   = 999012,
        DeployableMaterialsId = 999013,
        HousingItemsId = 999014,
        UhousingItemsId = 999015,
        BouquetsId = 999016;

    private readonly Dictionary<uint, Func<ItemRow, bool>> categoryFilters = new()
    {
        [RareGearId]          = IsRareGear,
        [WhiteGearId]         = IsWhiteGear,
        [MateriaId]           = IsMateria,
        [ConsumablesId]       = IsConsumable,
        [RegisterableId]      = IsRegisterableMiscellany,
        [GearBoxId]           = IsGearCoffer,
        [CatalystId]          = IsCatalyst,
        [CraftingMaterialsId] = IsCraftingMaterial,
        [CurrencyOtherId]     = IsCurrencyOtherMaterial,
        [TreasureMaterialsId] = IsTreasureMaterial,
        [SubmersiblePartsId]  = IsSubmersiblePart,
        [SchematicBoardId]    = IsSchematicBoardMaterial,
        [DeployableMaterialsId]= IsDeployableMaterial,
        [HousingItemsId]      = IsHousingItem,
        [UhousingItemsId]     = IsUntradableHousingItem,
        [BouquetsId]          = IsBouquetItem,
    };
    private static readonly HashSet<uint> materiaCategoryIds   = [57];
    private static readonly HashSet<uint> materialsCategoryIds = [44, 47, 48, 49, 50, 51, 52, 53, 54, 55, 58, 59];
    private static readonly HashSet<uint> submersiblePartsCategoryIds = [79];
    private static readonly HashSet<uint> housingItemsCategoryIds = [56, 65, 66, 67, 68, 69, 70, 71, 72];
    private static readonly GameStructs.InventoryType[] RetainerPages =
    {
        GameStructs.InventoryType.RetainerPage1,
        GameStructs.InventoryType.RetainerPage2,
        GameStructs.InventoryType.RetainerPage3,
        GameStructs.InventoryType.RetainerPage4,
        GameStructs.InventoryType.RetainerPage5,
        GameStructs.InventoryType.RetainerPage6,
        GameStructs.InventoryType.RetainerPage7,
    };
    private static readonly GameStructs.InventoryType[] PlayerInventoryPages =
    {
        GameStructs.InventoryType.Inventory1,
        GameStructs.InventoryType.Inventory2,
        GameStructs.InventoryType.Inventory3,
        GameStructs.InventoryType.Inventory4,
    };
    private static readonly GameStructs.InventoryType[] PlayerInventoryArmoryPages =
    {
        GameStructs.InventoryType.Inventory1,
        GameStructs.InventoryType.Inventory2,
        GameStructs.InventoryType.Inventory3,
        GameStructs.InventoryType.Inventory4,
        GameStructs.InventoryType.ArmoryBody,
        GameStructs.InventoryType.ArmoryEar,
        GameStructs.InventoryType.ArmoryFeets,
        GameStructs.InventoryType.ArmoryHands,
        GameStructs.InventoryType.ArmoryHead,
        GameStructs.InventoryType.ArmoryLegs,
        GameStructs.InventoryType.ArmoryMainHand,
        GameStructs.InventoryType.ArmoryNeck,
        GameStructs.InventoryType.ArmoryOffHand,
        GameStructs.InventoryType.ArmoryRings,
        GameStructs.InventoryType.ArmoryWaist,
        GameStructs.InventoryType.ArmoryWrist,        
    };    
    
#endregion

#region Types
    private sealed class TransferState
    {
        public volatile bool Running;
        public volatile int Moved;
        public volatile int Total;
    }
    private readonly record struct TransferPreview(
        int totalStacks,
        int transferStacks,
        int inventoryFreeSlots,
        int itemsToMove
    );
#endregion

#region Fields
    private static ExcelSheet<ItemRow>? itemSheetCache;
    private bool isFilterPanelVisible = false;
    private readonly HashSet<uint> selectedCategoryIds = new();
    private readonly Dictionary<uint, int> retainerCategoryCounts = new();

    private readonly TransferState transferSession = new();
    private CancellationTokenSource? cancellationTokenSource;

    private int transferDelayMsActive = 0;
    private int currentRetainerPageIndex = 0;
    private int currentRetainerSlotIndex = 0;
    private int currentInventoryPageIndex = 0;
    private int currentInventorySlotIndex = 0;
    private bool noInverse = true;
    private HashSet<uint>? seenUniqueDuringRun;
    private DateTime nextMoveAtUtc = DateTime.MinValue;

    private static readonly Random delayRandom = new();
    private int inventoryContainerOffset = 0;
    private int inventorySlotIndex = 0;
#endregion

#region Data Access & Timing
    /**
     * * Retrieves an item’s data row from the Excel sheet using its ID
     * <param name="itemId">The unique ID of the item to retrieve</param>
     * <return type="ItemRow?">The matching ItemRow, or null if not found</return>
    */
    private static ItemRow? GetItemRowById(uint itemId)
    {
        if (itemSheetCache == null)
            itemSheetCache = Plugin.DataManager.GetExcelSheet<ItemRow>(ClientLanguage.English);

        return itemSheetCache?.GetRow(itemId);
    }

    /**
     * * Resets the internal inventory navigation pointers
     * Used before starting a new transfer or inventory scan
     */
    private void ResetInventoryCursor()
    {
        inventoryContainerOffset = 0;
        inventorySlotIndex = 0;
    }

    /**
     * * Produces a randomized delay value around a given base delay
     * <param name="baseDelay">The base delay in milliseconds</param>
     * <return type="int">A humanized delay value in milliseconds</return>
     */
    private static int GenerateHumanizedDelay(int baseDelay)
    {
        if (baseDelay <= 0)
            return 0;

        int randomRange = Math.Max(20, (int)(baseDelay * 0.25));

        int delayOffset;
        lock (delayRandom) 
        {
            delayOffset = delayRandom.Next(-randomRange, randomRange + 1);
        }

        return Math.Max(20, baseDelay + delayOffset);
    }
#endregion

#region Category Filters

    /**
     * * Determines whether the given item is a white gear
     * <param name="item">The item row to evaluate</param>
     * <return type="bool">True if the item is equippable gear; otherwise, false</return>
     */ 
    private static bool IsWhiteGear(ItemRow item)
    {
        bool isGear = item.EquipSlotCategory.RowId > 0;
        bool isWhite = item.Rarity < 2;
        return isGear && isWhite;
    }

    /**
     * * Determines whether the given item is rare gear
     * <param name="item">The item row to evaluate</param>
     * <return type="bool">True if the item is gear and has rarity above normal</return>
     */
    private static bool IsRareGear(ItemRow item)
    {
        bool isGear = item.EquipSlotCategory.RowId > 0;
        bool isRare = item.Rarity > 1;

        return isGear && isRare;
    }

    /**
     * * Determines whether the given item is classified as materia
     * <param name="item">The item row to evaluate</param>
     * <return type="bool">True if the item belongs to a known materia search category</return>
     */
    private static bool IsMateria(ItemRow item)
    {
        bool isMateria = materiaCategoryIds.Contains(item.ItemSearchCategory.RowId);
        return isMateria;
    }

    /**
     * * Determines whether the given item is a consumable
     * <param name="item">The item row to evaluate</param>
     * <return type="bool">True if the item is non-gear and has an associated action</return>
     */
    private static bool IsConsumable(ItemRow item)
    {
        bool isNotGear = item.EquipSlotCategory.RowId == 0;
        bool hasAction = item.ItemAction.RowId > 0;
        bool isGysahl = item.RowId == 4868;
        bool isAetheryteTicket = item.RowId >= 21069 && item.RowId <= 21071;
        //might want to add a checkbox somewhere
        return isNotGear && hasAction && !isGysahl && !isAetheryteTicket && !IsGearCoffer(item) && !IsRegisterableMiscellany(item);
    }

    /**
     * * Determines whether the given item is a registerable miscellany
     * <param name="item">The item row to evaluate</param>
     * <return type="bool">True if the item is non-gear and has an associated action</return>
     */
    private static bool IsRegisterableMiscellany(ItemRow item)
    {
        //50435 searchcategory 90
        //43573 searchcategory 75
        bool isNotGear = item.EquipSlotCategory.RowId == 0;
        bool hasAction = item.ItemAction.RowId > 0;
        bool consumable = isNotGear && hasAction;

        bool registerable = item.ItemSearchCategory.RowId == 90;
        bool minion = item.ItemSearchCategory.RowId == 75;

        return consumable && (registerable || minion);
    }      

    /**
     * * Determines whether the given item is a gear coffer
     * <param name="item">The item row to evaluate</param>
     * <return type="bool">True if the item is non-gear and has an associated action</return>
     */    
    private static bool IsGearCoffer(ItemRow item)
    {
        bool isgearcoffer = item.ItemUICategory.RowId == 61; //Miscellany
        //bool hascoffer = item.Name.ToString().Contains("Coffer");
        //bool isntventurecoffer = !item.Name.ToString().Contains("Venture");
        //item.ItemAction.Value.Action.RowId
        bool cantbesold = item.ItemSearchCategory.RowId == 0;
        bool chestactiontype = item.ItemAction.RowId == 388 || item.ItemAction.RowId == 1085;
        bool cantstack = item.StackSize == 1;
        return isgearcoffer && cantbesold && chestactiontype && cantstack;
    }

    /**
     * * Determines whether the given item is a catalyst. (glamour prism, dispeller, graded dark matter, dark matter cluster)
     * <param name="item">The item row to evaluate</param>
     * <return type="bool">True if the item is in materials categories and has no action</return>
     */
    private static bool IsCatalyst(ItemRow item)
    {
        bool isCatalystUiCategory = item.ItemUICategory.RowId == 60;
        bool isGlamourPrism = item.RowId == 21800;
        return 
        (
            isCatalystUiCategory && 
            !isGlamourPrism
        );
    }

    /**
     * * Determines whether the given item is a crafting material
     * <param name="item">The item row to evaluate</param>
     * <return type="bool">True if the item is in materials categories and has no action</return>
     */
    private static bool IsCraftingMaterial(ItemRow item)
    {
        bool isMaterial = materialsCategoryIds.Contains(item.ItemSearchCategory.RowId);
        bool isSeaFood = item.ItemSearchCategory.RowId == 46;
        bool isCMaterial = isMaterial || isSeaFood;
        bool hasNoAction = item.ItemAction.RowId == 0;
        bool isGlamourPrism = item.RowId == 21800;
        bool isMarineWax = item.RowId == 23222; //has itemsearch category 60 (miscellany)
        //uicategory 56 => parts like 5580 "vanya silk crakow lining"   
        //uicategory 60 => catalyst ? (glamour prism, darkmatter) / searchcategory 59
        //using FFXIVClientStructs.FFXIV.Client.UI.Agent; //from FCCH MoveManager.cs
        //AgentInventoryContext.Instance()->LowerItemQuality(item, inventorypage, slot, 0);
        bool regularCraftingMat =             
        (
            isCMaterial && 
            hasNoAction && 
            !isGlamourPrism && //redundant with catalyst
            !IsCatalyst(item) &&
            !IsSchematicBoardMaterial(item) && 
            !IsDeployableMaterial(item) && 
            !IsTreasureMaterial(item) &&
            !IsCurrencyOtherMaterial(item)
        );

        return regularCraftingMat || isMarineWax;
    }
    
    /**
     * * Determines whether the given item is a currency bought crafting material
     * <param name="item">The item row to evaluate</param>
     * <return type="bool">True if the item is in materials categories and has no action</return>
     */    
    private static bool IsCurrencyOtherMaterial(ItemRow item)
    {
        bool craftMaterialsCurrencyOther = (
            //(item.RowId == ) || //
            //not really treasure materials but cost scrip or tomestone, and are used for white gear glamour craft
            (item.RowId == 45987) || //cronopio skin - 20 tomestone of mathematics
            (item.RowId == 39595) || //gripgel - 250 purple crafting scrip
            (item.RowId == 7775) || //glazenut - gardening
            (item.RowId == 13754) || //blood pepper - gardening
            (item.RowId == 20791) || //cloudsbreath - gardening
            (item.RowId == 17574) || //flannel - 200 purple gathering or crafting scrip
            (item.RowId == 13744) || //wind silk - 75 purple gathering scrip
            (item.RowId == 19992) || //doman weave - 10 poetics
            (item.RowId == 13748) || //cloud cloth - 20 purple gathering scrip
            (item.RowId == 32937) || //saffron cloth - 40 poetics
            (item.RowId == 32947) || //gardenia fruit - 20 poetics
            (item.RowId == 45978) || //diatryma felt - 40 tomestone of mathematics
            (item.RowId == 37832)  //titanoboa leather - 20 poetics
        );

        return craftMaterialsCurrencyOther;
    }

    /**
     * * Determines whether the given item is a crafting material from treasure hunt (also gotten from quick venture) specifically those for glamour crafts
     * <param name="item">The item row to evaluate</param>
     * <return type="bool">True if the item is in materials categories and has no action</return>
     */
    private static bool IsTreasureMaterial(ItemRow item)
    {
        bool craftMaterialsFromTreasureHunt = (
            //(item.RowId == ) || //
            //HW treasure dungeon
            (item.RowId == 16907) || //new world macrame
            (item.RowId == 16906) || //wool top
            //STB treasure dungeon
            (item.RowId == 20271) || //rathanhia leaves
            (item.RowId == 21205) || //hannish wool
            (item.RowId == 21206) || //thavnairian wool
            (item.RowId == 24336) || //whisperfine fleece
            (item.RowId == 20270) || //nagxian silk
            (item.RowId == 22464) || //ala mhigan weave
            (item.RowId == 23375) || //cloth cointe quaint
            //SHB treasure dungeon
            (item.RowId == 33869) || //frontier cloth
            (item.RowId == 27977) || //enchanted elm lumber
            //EW treasure dungeon
            (item.RowId == 39589) || //hagnos cloth
            (item.RowId == 41653) || //athletikos cloth
            (item.RowId == 36628) || //cloth of happiness
            //DT treasure dungeon
            (item.RowId == 48229) || //flexible leather band
            (item.RowId == 45072) || //corduroy felt
            (item.RowId == 46261) || //summerlight linen
            (item.RowId == 50830) || //cloth of legend
            (item.RowId == 52631) || //indigo twill cloth
            (item.RowId == 52630) || //sweatcloth
            //from treasure map
            (item.RowId == 8026) || //mist silk
            (item.RowId == 7772) || //mariner cotton cloth
            (item.RowId == 12646) || //thavnairian silk
            (item.RowId == 12645) || //thavnairian leather
            (item.RowId == 23376) || //leather cointe quaint
            (item.RowId == 22463) || //holy linen
            (item.RowId == 20273) || //nagxian leather
            (item.RowId == 20272) || //taoist wool
            (item.RowId == 24804) || //waterproof cloth
            (item.RowId == 28905) || //shell leather
            (item.RowId == 27975) || //indigo cloth
            (item.RowId == 27974) || //steer hide
            (item.RowId == 27953) || //fine glass fiber
            (item.RowId == 27976) || //fireglass leather
            (item.RowId == 7044) || //exquisite buttons
            (item.RowId == 7043) || //ahriman tears
            (item.RowId == 7032) || //lotus leaf
            (item.RowId == 44133) || //mamonite
            (item.RowId == 44132) || //timeless turali cloth
            (item.RowId == 40700) || //malake karpasos
            (item.RowId == 38686) || //exciting leather
            (item.RowId == 7042) || //chinchilla hide
            (item.RowId == 7771) || //dress material
            (item.RowId == 7047) || //frosted glass lens
            (item.RowId == 36629) || //phonograph plate
            (item.RowId == 7045) || //morbol saliva
            (item.RowId == 8021) || //rush grass
            (item.RowId == 7046) || //cutting oil
            (item.RowId == 14931) || //antique rug
            (item.RowId == 14932) || //antique leather
            (item.RowId == 13746) || //falconer hemp
            (item.RowId == 7041)  || //blue fox hide
            (item.RowId == 30135) || //cloth softening powder
            (item.RowId == 31091) || //sea cloth
            (item.RowId == 33027)    //calf leather
        );

        return craftMaterialsFromTreasureHunt;
    }

    /**
     * * Determines whether the given item is a schematic board material
     * <param name="item">The item row to evaluate</param>
     * <return type="bool">True if the item is in materials categories and has no action</return>
     */ 
    private static bool IsSchematicBoardMaterial(ItemRow item)
    {
        bool isBoardMaterial = (
            (item.RowId == 37083) || //basaltic clay
            (item.RowId == 12920) || //red clay
            (item.RowId == 12921) || //black clay
            (item.RowId == 29590) || //pelagic clay
            (item.RowId == 26528) || //cocobolo lumber
            (item.RowId == 12929) || //empty cluster
            (item.RowId == 12928) || //empty crystal
            (item.RowId == 12927) || //aetherochemical fiber
            (item.RowId == 12926) || //synthetic fiber
            (item.RowId == 12925) || //balsa wood lumber
            (item.RowId == 12924) || //balsa wood scrap
            (item.RowId == 12923) || //bamboo weave
            (item.RowId == 12922)    //synthetic resin
        );

        return isBoardMaterial;
    }

    /**
     * * Determines whether the given item is a deployable material
     * <param name="item">The item row to evaluate</param>
     * <return type="bool">True if the item is in materials categories and has no action</return>
     */ 
    private static bool IsDeployableMaterial(ItemRow item)
    {
        bool DeployableMaterial = (
            //(item.RowId == ) || //
            (item.RowId == 10099) || //unaspected crystal
            (item.RowId == 12230) || //deep-blue cluster
            (item.RowId == 12229) || //deep-blue crystal
            (item.RowId == 12228) || //deep-green cluster
            (item.RowId == 12227) || //deep-green crystal
            (item.RowId == 12226) || //deep-red cluster
            (item.RowId == 12225) || //deep-red crystal            
            (item.RowId == 12231) || //farreach canvas
            (item.RowId == 12220) || //red moko grass
            (item.RowId == 13758) || //cassia block
            (item.RowId == 12232) || //cassia lumber
            (item.RowId == 12221) || //cassia log
            (item.RowId == 26532) || //pure titanium plate
            (item.RowId == 26531) || //pure titanium ore
            (item.RowId == 26530) || //cryptomeria lumber
            (item.RowId == 26529) || //cryptomeria log
            (item.RowId == 14935) || //galvanized garlond stel
            (item.RowId == 12913) || //garlond steel
            (item.RowId == 12223) || //raw celestine
            (item.RowId == 12912) || //celestine
            (item.RowId == 12224)    //vivianite
        );

        return DeployableMaterial;
    }                 

    /**
     * * Determines whether the given item is classified as submersible parts
     * <param name="item">The item row to evaluate</param>
     * <return type="bool">True if the item belongs to a known submersible parts search category</return>
     */
    private static bool IsSubmersiblePart(ItemRow item)
    {
        //airship part UICategory hull 90 sail 91 forecastle 93 aftcastle 92 
        //submarine part UICategory hull 101 bow 103 bridge 104 stern 102
        bool isSubmersiblePart = submersiblePartsCategoryIds.Contains(item.ItemSearchCategory.RowId);
        return isSubmersiblePart;
    }

    /**
     * * Determines whether the given item is a housing item
     * <param name="item">The item row to evaluate</param>
     * <return type="bool">True if the item belongs to a known submersible parts search category</return>
     */    
    private static bool IsHousingItem(ItemRow item)
    {
        //just noticed i was using ItemUICategory rather than ItemSearchCategory
        //Combed Wool Rug https://www.garlandtools.org/db/#item/15976
        //https://v2.xivapi.com/api/sheet/Item/15976?fields=Name,ItemUICategory.Name,ItemUICategory.RowId
        //roof 65 house 66 window 67 door 68 69 70 71 fence 72 wall 73 flooring 74 light 75 outdoor 76 furnishing 57 tabletop 78 wallmounted 79 rug 80
        //with ItemUICategory private static readonly HashSet<uint> housingItemsCategoryIds = [57, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80];
        //bool isHousing = housingItemsCategoryIds.Contains(item.ItemUICategory.RowId);


        //https://v2.xivapi.com/api/sheet/Item/15976?fields=Name,RowId,ItemSearchCategory.RowId 
        //itemsearchcategory exteriorfixture 65 interiorfixt 66 outdoor 67 furnishing 56 chair 68 table 69 tabletop 70 wallmounted 71 rug 72 
        //with ItemSearchCategory private static readonly HashSet<uint> housingItemsCategoryIds = [56, 65, 66, 67, 68, 69, 70, 71, 72];
        bool isHousing = housingItemsCategoryIds.Contains(item.ItemSearchCategory.RowId);                        
        return isHousing && !IsUntradableHousingItem(item);
    } 

    /**
     * * Determines whether the given item is a untradable housing item 
     * <param name="item">The item row to evaluate</param>
     * <return type="bool">True if the item belongs to a known submersible parts search category</return>
     */
    private static bool IsUntradableHousingItem(ItemRow item)
    {
        HashSet<uint> uhousingItemsCategoryIds = [57, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80];
        bool untradable = item.ItemSearchCategory.RowId == 0;
        bool isHousing = uhousingItemsCategoryIds.Contains(item.ItemUICategory.RowId);                        
        return isHousing && untradable;
    }          

    /**
     * * Determines whether the given item is a bouquet housing item 
     * <param name="item">The item row to evaluate</param>
     * <return type="bool">True if the item belongs to a known submersible parts search category</return>
     */
    private static bool IsBouquetItem(ItemRow item)
    {
        //Dictionary<ClientLanguage, string> BouquetWord = new()
        //{
            //[ClientLanguage.English]  = "Bouquets can be placed within select vases for display in estates.",
            //[ClientLanguage.French]   = "Peut être disposé dans un vase spécial",
            //[ClientLanguage.German]   = "Kann in spezielle Vasen gestellt werden.",
            //[ClientLanguage.Japanese] = "特定の花瓶に挿すことができる",
        //};
        HashSet<uint> uBouquetItemsCategoryIds = [61];
        //var word = BouquetWord[Plugin.DataManager.Language];
        bool tradable = item.ItemSearchCategory.RowId > 0;
        bool isBouquet = uBouquetItemsCategoryIds.Contains(item.ItemUICategory.RowId);
        bool describesBouquet = item.Description.ToString().Contains("Bouquets can be placed within select vases for display in estates.", StringComparison.OrdinalIgnoreCase);                        
        return isBouquet && tradable && describesBouquet;
    }          



#endregion

#region Filters & Counts
    /**
     * * Increases the stored count for the specified category by one
     * <param name="categoryId">The ID of the category to increment</param>
     */
    private void IncrementCategory(uint categoryId)
    {
        retainerCategoryCounts[categoryId] = retainerCategoryCounts.GetValueOrDefault(categoryId) + 1;
    }

    /**
     * * Clears and rebuilds category count data for the retainer’s inventory
     */
    private unsafe void RecountRetainerCategoryCounts()
    {
        GameStructs.InventoryType[] containerList;
        if (noInverse)
        {
            containerList = RetainerPages;
        } 
        else
        {
            containerList = PlayerInventoryPages;
        }

        retainerCategoryCounts.Clear();

        var inv = GameStructs.InventoryManager.Instance();
        if (inv == null)
            return;

        foreach (var page in containerList)
        {
            var container = inv->GetInventoryContainer(page);
            if (container == null)
                continue;

            for (int slotIndex = 0; slotIndex < container->Size; slotIndex++)
            {
                var slot = container->GetInventorySlot(slotIndex);
                if (slot == null || slot->ItemId == 0 || slot->Quantity == 0)
                    continue;

                var itemRow = GetItemRowById(slot->ItemId);
                if (itemRow == null)
                    continue;

                var row = itemRow.Value;

                foreach (var (categoryId, filter) in categoryFilters)
                {
                    if (filter(row))
                        IncrementCategory(categoryId);
                }
            }
        }            

    }

    /**
     * * Determines whether the given item passes the currently selected category filters
     * <param name="item">The item row to evaluate</param>
     * <return type="bool">True if the item matches at least one selected filter; otherwise, false</return>
     */
    private bool MatchesTransferFilters(ItemRow item)
    {
        if (selectedCategoryIds.Count == 0)
            return true;

        foreach (var categoryId in selectedCategoryIds)
        {
            if (!categoryFilters.TryGetValue(categoryId, out var filter))
                continue;

            if (filter(item))
                return true;
        }

        return false;
    }
#endregion

#region Inventory Analysis
    /**
     * * Scans the player’s inventory to count free slots and stack capacities
     * <param name="inv">Pointer to the InventoryManager instance</param>
     * <param name="stackSpace">Dictionary updated with remaining stack space per item ID</param>
     * <param name="inventoryItems">Set updated with all unique item IDs found in inventory</param>
     * <return type="int">The total number of free slots in the player’s inventory</return>
     */
    private unsafe int CountInventory(
        GameStructs.InventoryManager* inv, 
        Dictionary<uint, int> stackSpace, 
        HashSet<uint> inventoryItems)
    {
        int freeSlots = 0;
        GameStructs.InventoryType[] containerList;
        if (noInverse)
        {
            containerList = PlayerInventoryPages;
        } 
        else
        {
            containerList = RetainerPages;
        }

        foreach (var page in containerList)
        {
            var container = inv->GetInventoryContainer(page);
            if (container == null)
                continue;

            for (int s = 0; s < container->Size; s++)
            {
                var slot = container->GetInventorySlot(s);
                if (slot == null || slot->ItemId == 0)
                {
                    freeSlots++;
                    continue;
                }

                inventoryItems.Add(slot->ItemId);

                var row = GetItemRowById(slot->ItemId);
                if (row == null)
                    continue;

                int maxStack = (int)row.Value.StackSize;
                if (maxStack <= 1 || slot->Quantity >= maxStack)
                    continue;

                int remaining = Math.Max(0, maxStack - slot->Quantity);
                if (remaining > 0)
                    stackSpace[slot->ItemId] =
                        stackSpace.GetValueOrDefault(slot->ItemId) + remaining;
            }
        }

        return freeSlots;
    }

    /**
     * * Scans all retainer inventory pages and counts item stacks
     * <param name="inv">Pointer to the InventoryManager instance</param>
     * <param name="items">Set of item IDs currently in the player’s inventory</param>
     * <param name="moveStacks">Dictionary mapping item IDs to their stack quantities for transfer</param>
     * <param name="totalStacks">Reference counter incremented for every stack found</param>
     * <param name="transferStacks">Reference counter incremented for stacks eligible for transfer</param>
     */
    private unsafe void CountRetainerStacks(
        GameStructs.InventoryManager* inv,
        HashSet<uint> items,
        Dictionary<uint, List<int>> moveStacks,
        ref int totalStacks,
        ref int transferStacks)
    {
        var seenUniqueFromRetainer = new HashSet<uint>();
        GameStructs.InventoryType[] containerList;
        
        if (noInverse)
        {
            containerList = RetainerPages;
        } 
        else
        {
            containerList = PlayerInventoryPages;
        }

        foreach (var page in containerList)
        {
            var container = inv->GetInventoryContainer(page);
            if (container == null)
                continue;

            for (int s = 0; s < container->Size; s++)
            {
                var slot = container->GetInventorySlot(s);
                if (slot == null || slot->ItemId == 0 || slot->Quantity == 0)
                    continue;

                totalStacks++;

                var row = GetItemRowById(slot->ItemId);
                if (row == null || !MatchesTransferFilters(row.Value))
                    continue;

                if (row.Value.IsUnique)
                {
                    if (items.Contains(slot->ItemId))
                        continue;
                    if (!seenUniqueFromRetainer.Add(slot->ItemId))
                        continue;
                }

                transferStacks++;

                if (!moveStacks.TryGetValue(slot->ItemId, out var list))
                {
                    list = new List<int>();
                    moveStacks[slot->ItemId] = list;
                }
                list.Add(slot->Quantity);
            }
        }
        //Plugin.ChatGui.Print("[MassWithdraw] Transfer stacks: …");
        //Plugin.Log.Debug($"[MassWithdraw] Total stacks: {totalStacks}");
        //Plugin.Log.Debug($"[MassWithdraw] Transfer stacks before mergeable stack: {transferStacks}");
    }

    /**
     * * Calculates how many retainer stacks can be merged into existing inventory stacks
     * <param name="moveStacks">Dictionary of item IDs mapped to their retainer stack quantities</param>
     * <param name="stackSpace">Dictionary of item IDs with available stacking capacity in inventory</param>
     * <return type="int">The total number of stacks that can be merged into existing ones</return>
     */
    private static int CountMergeableStacks(
        Dictionary<uint, List<int>> moveStacks,
        Dictionary<uint, int> stackSpace)
    {
        int mergeable = 0;

        foreach (var (itemId, stacks) in moveStacks)
        {
            int remainingCap = stackSpace.GetValueOrDefault(itemId);
            if (remainingCap <= 0)
                continue;

            stacks.Sort();

            foreach (var qty in stacks)
            {
                if (qty <= remainingCap)
                {
                    mergeable++;
                    remainingCap -= qty;
                }
                else
                {
                    break;
                }
            }
        }

        return mergeable;
    }

    /**
     * * Checks if a specific item exists in the player’s inventory
     * <param name="itemId">The unique ID of the item to check for</param>
     * <return type="bool">True if the item is present in the player’s inventory; otherwise, false</return>
     */
    //todo merge hasitemininventory and hasiteminretainer in hasitemintargetinventory
    private static unsafe bool HasItemInInventory(uint itemId)
    {
        if (itemId == 0)
            return false;

        var inv = GameStructs.InventoryManager.Instance();
        if (inv == null)
            return false;

        foreach (var page in PlayerInventoryPages)
        {
            var container = inv->GetInventoryContainer(page);
            if (container == null)
                continue;

            for (int slot = 0; slot < container->Size; slot++)
            {
                var currentSlot = container->GetInventorySlot(slot);
                if (currentSlot == null)
                    continue;

                bool isSameItem = currentSlot->ItemId == itemId;
                bool hasQuantity = currentSlot->Quantity > 0;

                if (isSameItem && hasQuantity)
                    return true;
            }
        }

        return false;
    }

    private static unsafe bool HasItemInRetainer(uint itemId)
    {
        if (itemId == 0)
            return false;

        var inv = GameStructs.InventoryManager.Instance();
        if (inv == null)
            return false;

        foreach (var page in RetainerPages)
        {
            var container = inv->GetInventoryContainer(page);
            if (container == null)
                continue;

            for (int slot = 0; slot < container->Size; slot++)
            {
                var currentSlot = container->GetInventorySlot(slot);
                if (currentSlot == null)
                    continue;

                bool isSameItem = currentSlot->ItemId == itemId;
                bool hasQuantity = currentSlot->Quantity > 0;

                if (isSameItem && hasQuantity)
                    return true;
            }
        }

        return false;
    }    
#endregion

#region Slot Selection
    /**
     * * Searches the player’s inventory for an existing partial stack of the specified item
     * <param name="itemId">The ID of the item to find a stackable slot for</param>
     * <param name="targetContainer">Outputs the inventory container where stacking is possible</param>
     * <param name="targetSlot">Outputs the slot index of the stackable item</param>
     * <return type="bool">True if a valid stackable slot is found; otherwise, false</return>
     */
    private unsafe bool FindStackableSlot(uint itemId, out GameStructs.InventoryType targetContainer, out int targetSlot, bool retainerToInv, bool isHQ)
    {
        GameStructs.InventoryType[] currentPagesList = [];
        if (retainerToInv) {
            targetContainer = default;
            currentPagesList = PlayerInventoryPages;
        } else {
            targetContainer = GameStructs.InventoryType.RetainerPage1;
            currentPagesList = RetainerPages;
        }
        targetSlot = -1;

        if (itemId == 0)
            return false;

        var inv = GameStructs.InventoryManager.Instance();
        if (inv == null)
            return false;

        var row = GetItemRowById(itemId);
        if (row == null)
            return false;

        int maxStack = (int)row.Value.StackSize;
        if (maxStack <= 1)
            return false;

        foreach (var page in currentPagesList)
        {
            var container = inv->GetInventoryContainer(page);
            if (container == null)
                continue;

            for (int slot = 0; slot < container->Size; slot++)
            {
                var s = container->GetInventorySlot(slot);
                bool targetHq = s->Flags.HasFlag(InventoryItem.ItemFlags.HighQuality);
                if (s == null)
                    continue;

                if ((targetHq && !isHQ) || (!targetHq && isHQ))
                    continue;
                
                if (s->ItemId == itemId && s->Quantity > 0 && s->Quantity < maxStack)
                {
                    targetContainer = page;
                    targetSlot = slot;
                    return true;
                }
            }
        }
        
        return false;
    }

    /**
     * * Searches the player’s inventory for the next available empty slot
     * <param name="targetContainer">Outputs the container type where a free slot was found</param>
     * <param name="targetSlot">Outputs the index of the available inventory slot</param>
     * <return type="bool">True if a free slot is found; otherwise, false</return>
     */
    private unsafe bool FindFreeBagSlot(out GameStructs.InventoryType targetContainer, out int targetSlot, bool retainerToInv)
    {
        int bagCount;
        if (retainerToInv) {
            targetContainer = default;
            bagCount = 4;
        } else {
            targetContainer = GameStructs.InventoryType.RetainerPage1;
            bagCount = 7;
        }
        targetSlot = -1;

        var inv = GameStructs.InventoryManager.Instance();
        if (inv == null)
            return false;

        //const int bagCount = 4;
        

        for (int checkedContainers = 0; checkedContainers < bagCount; checkedContainers++)
        {
            GameStructs.InventoryType containerType;
            if (retainerToInv) {
                containerType = PlayerInventoryPages[inventoryContainerOffset];
            } else {
                containerType = RetainerPages[inventoryContainerOffset];
            }
            var container     = inv->GetInventoryContainer(containerType);
            int size          = container != null ? container->Size : 0;

            if (size > 0 && container != null)
            {
                int startSlot = inventorySlotIndex;

                for (int s = startSlot; s < size; s++)
                {
                    var slot = container->GetInventorySlot(s);
                    if (slot == null || slot->ItemId == 0)
                    {
                        targetContainer = containerType;
                        targetSlot      = s;

                        if (s + 1 < size)
                        {
                            inventorySlotIndex = s + 1;
                        }
                        else
                        {
                            inventorySlotIndex = 0;
                            inventoryContainerOffset = (inventoryContainerOffset + 1) % bagCount;
                        }

                        return true;
                    }
                }
            }

            inventorySlotIndex = 0;
            inventoryContainerOffset = (inventoryContainerOffset + 1) % bagCount;
        }

        return false;
    }
#endregion

#region Transfer
    /**
     * * Computes a summary of a potential transfer from retainer to inventory
     * <return type="object">A summary containing total stacks, eligible stacks, free inventory slots, and movable items</return>
     */
    private unsafe TransferPreview GenerateTransferPreview()
    {
        var inv = GameStructs.InventoryManager.Instance();
        if (inv == null)
            return new TransferPreview(0, 0, 0, 0);

        var stackSpace = new Dictionary<uint, int>();
        var inventoryItems = new HashSet<uint>();
        int inventoryFreeSlots = CountInventory(inv, stackSpace, inventoryItems);
        //Plugin.ChatGui.Print($"CountInventory = {inventoryFreeSlots}");
        var moveStacks = new Dictionary<uint, List<int>>();
        int totalStacks = 0;
        int transferStacks = 0;
        CountRetainerStacks(inv, inventoryItems, moveStacks, ref totalStacks, ref transferStacks);
        //Plugin.ChatGui.Print($"transferStacks = {transferStacks}");

        int mergeable = CountMergeableStacks(moveStacks, stackSpace);
        //Plugin.ChatGui.Print($"mergeable = {mergeable}");

        int unmergedStacks = Math.Max(0, transferStacks - mergeable);
        //Plugin.ChatGui.Print($"unmergedStacks = {unmergedStacks}");
        int itemsToMove = mergeable + Math.Min(unmergedStacks, inventoryFreeSlots);
        //Plugin.ChatGui.Print($"itemsToMove = {itemsToMove}");

        return new TransferPreview(totalStacks, transferStacks, inventoryFreeSlots, itemsToMove);
    }

    /**
     * * Starts an item transfer session if one is not already running
     * <param name="transferDelayMs">Base delay (ms) between moves for throttling</param>
     */
    private void StartTransfer(int transferDelayMs)
    {
        if (transferSession.Running)
            return;

        var preview = GenerateTransferPreview();
        if (preview.itemsToMove <= 0)
        {
            Plugin.ChatGui.Print("[MassWithdraw] Nothing to transfer.");
            return;
        }

        transferSession.Moved = 0;
        transferSession.Total = preview.itemsToMove;
        transferSession.Running = true;

        ResetInventoryCursor();

        cancellationTokenSource = new CancellationTokenSource();
        transferDelayMsActive = transferDelayMs;
        currentRetainerPageIndex = 0;
        currentRetainerSlotIndex = 0;
        currentInventoryPageIndex = 0;
        currentInventorySlotIndex = 0;
        seenUniqueDuringRun = new HashSet<uint>();
        nextMoveAtUtc = DateTime.UtcNow;
    }

    private void OnFrameworkUpdate(IFramework _)
    {
        if (!transferSession.Running)
            return;

        if (cancellationTokenSource == null || cancellationTokenSource.IsCancellationRequested)
        {
            StopTransfer($"[MassWithdraw] Cancelled. Moved so far: {transferSession.Moved} item(s).");
            return;
        }

        if (!IsRetainerUIOpen())
        {
            StopTransfer($"[MassWithdraw] Stopped: retainer closed. Moved {transferSession.Moved} item(s).");
            return;
        }

        var now = DateTime.UtcNow;
        if (now < nextMoveAtUtc)
            return;

        bool didMove;
        if (noInverse) {
            didMove = TryMoveOneItem();
        } else {
            didMove = TryMoveOneItemReverse();
        }
        //bool didMove = TryMoveOneItem();
        //bool didMove = TryMoveOneItemReverse();

        if (!transferSession.Running)
            return;

        if (didMove)
        {
            System.Threading.Interlocked.Increment(ref transferSession.Moved);

            int delay = GenerateHumanizedDelay(transferDelayMsActive);

            int roll;
            lock (delayRandom)
                roll = delayRandom.Next(0, 100);

            if (roll < 7)
            {
                int extra;
                lock (delayRandom)
                    extra = delayRandom.Next(100, 250);

                delay += GenerateHumanizedDelay(extra);
            }

            // small breather every 10 moves (helps FPS)
            if (transferSession.Moved % 10 == 0)
                delay += 250;

            nextMoveAtUtc = now.AddMilliseconds(Math.Max(20, delay));
        }
        else
        {
            nextMoveAtUtc = now;
        }
    }

    private void StopTransfer(string message)
    {
        transferSession.Running = false;

        var cts = cancellationTokenSource;
        cancellationTokenSource = null;
        cts?.Dispose();

        seenUniqueDuringRun?.Clear();
        seenUniqueDuringRun = null;

        Plugin.ChatGui.Print(message);
    }
    //todo merge trymoveoneitem and trymoveoneitemreverse in trymoveoneitemtowardtarget
    private unsafe bool TryMoveOneItem()
    {
        var inv = GameStructs.InventoryManager.Instance();
        if (inv == null)
        {
            StopTransfer("[MassWithdraw] InventoryManager not available.");
            return false;
        }

        seenUniqueDuringRun ??= new HashSet<uint>();

        while (currentRetainerPageIndex < RetainerPages.Length)
        {
            var page = RetainerPages[currentRetainerPageIndex];
            var container = inv->GetInventoryContainer(page);
            if (container == null)
            {
                currentRetainerPageIndex++;
                currentRetainerSlotIndex = 0;
                continue;
            }

            int pageSize = container->Size;

            while (currentRetainerSlotIndex < pageSize)
            {
                int slot = currentRetainerSlotIndex++;
                var it = container->GetInventorySlot(slot);
                if (it == null || it->ItemId == 0 || it->Quantity == 0)
                    continue;

                uint itemId = it->ItemId;
                bool originHQ = it->Flags.HasFlag(InventoryItem.ItemFlags.HighQuality);
                int quantity = it->Quantity;

                var row = GetItemRowById(itemId);
                if (row == null || !MatchesTransferFilters(row.Value))
                    continue;

                if (row.Value.IsUnique)
                {
                    if (!seenUniqueDuringRun.Add(itemId))
                        continue;
                    if (HasItemInInventory(itemId))
                        continue;
                }

                GameStructs.InventoryType targetContainer;
                int targetSlot;

                if (!FindStackableSlot(itemId, out targetContainer, out targetSlot, true,originHQ) &&
                    !FindFreeBagSlot(out targetContainer, out targetSlot, true))
                {
                    StopTransfer($"[MassWithdraw] Stopped: no free bag space. Moved {transferSession.Moved} item(s).");
                    return false;
                }

                var currentSlot = container->GetInventorySlot(slot);
                if (currentSlot == null || currentSlot->ItemId != itemId || currentSlot->Quantity != quantity)
                    return false;

                _ = inv->MoveItemSlot(page, (ushort)slot, targetContainer, (ushort)targetSlot, true);
                return true;
            }

            currentRetainerPageIndex++;
            currentRetainerSlotIndex = 0;
        }

        StopTransfer($"[MassWithdraw] Done. Moved total: {transferSession.Moved} item(s).");
        return false;
    }

    private unsafe bool TryMoveOneItemReverse()
    {
        var inv = GameStructs.InventoryManager.Instance();
        if (inv == null)
        {
            StopTransfer("[MassWithdraw] InventoryManager not available.");
            return false;
        }

        seenUniqueDuringRun ??= new HashSet<uint>();
        //Plugin.ChatGui.Print($"currentInventoryPageIndex = {currentInventoryPageIndex}");
        //Plugin.ChatGui.Print($"PlayerInventoryPages.Length = {PlayerInventoryPages.Length}");
        while (currentInventoryPageIndex < PlayerInventoryPages.Length)
        {
            var page = PlayerInventoryPages[currentInventoryPageIndex];
            var container = inv->GetInventoryContainer(page);
            if (container == null)
            {
                currentInventoryPageIndex++;
                currentInventorySlotIndex = 0;
                continue;
            }

            int pageSize = container->Size;

            while (currentInventorySlotIndex < pageSize)
            {
                if (currentInventorySlotIndex > 0)
                {
                    var tempinv = container->GetInventorySlot(currentInventorySlotIndex-1);
                    if (!(tempinv == null || tempinv->ItemId == 0 || tempinv->Quantity == 0))
                    {
                        currentInventorySlotIndex--;
                    }
                }
                int slot = currentInventorySlotIndex++; //need to check whether slot is empty becore incrementing (case where stackable was found)
                var it = container->GetInventorySlot(slot);
                if (it == null || it->ItemId == 0 || it->Quantity == 0)
                    continue;

                uint itemId = it->ItemId;
                bool originHQ = it->Flags.HasFlag(InventoryItem.ItemFlags.HighQuality);
                int quantity = it->Quantity;

                var row = GetItemRowById(itemId);
                if (row == null || !MatchesTransferFilters(row.Value))
                    continue;

                if (row.Value.IsUnique)
                {
                    if (!seenUniqueDuringRun.Add(itemId))
                        continue;
                    if (HasItemInRetainer(itemId))
                        continue;
                }

                GameStructs.InventoryType targetContainer;
                int targetSlot;

                if (!FindStackableSlot(itemId, out targetContainer, out targetSlot, false,originHQ) &&
                    !FindFreeBagSlot(out targetContainer, out targetSlot, false))
                {
                    StopTransfer($"[MassWithdraw] Stopped: no free bag space. Moved {transferSession.Moved} item(s).");
                    return false;
                }

                var currentSlot = container->GetInventorySlot(slot);
                if (currentSlot == null || currentSlot->ItemId != itemId || currentSlot->Quantity != quantity)
                    return false;

                _ = inv->MoveItemSlot(page, (ushort)slot, targetContainer, (ushort)targetSlot, true);
                return true;
            }

            currentInventoryPageIndex++;
            currentInventorySlotIndex = 0;
        }

        StopTransfer($"[MassWithdraw] Done. Moved total: {transferSession.Moved} item(s).");
        return false;
    }       

#endregion

#region Commands
    /**
     * * Entry point triggered by the user command to begin item transfer
     */
    public void StartTransferFromCommand()
    {
        if (!IsRetainerUIOpen())
        {
            Plugin.ChatGui.PrintError("[MassWithdraw] Open your Retainer’s inventory window first.");
            return;
        }

        var preview = GenerateTransferPreview();

        if (preview.itemsToMove <= 0)
        {
            var msg =
                preview.totalStacks == 0        ? "Retainer inventory is empty."
              : preview.transferStacks == 0     ? "No items match the selected filters."
              : preview.inventoryFreeSlots == 0 ? "Inventory full."
                                                : "Nothing to transfer.";

            Plugin.ChatGui.PrintError($"[MassWithdraw] {msg}");
            return;
        }

        StartTransfer(transferDelayMs);
        Plugin.ChatGui.Print("[MassWithdraw] Transfer started…");
    }
#endregion
}
