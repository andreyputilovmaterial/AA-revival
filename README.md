# AA-revival
New version of AA (2025)

## How to use it ##
Go to
[Releases](../../releases/latest)
page and downloads all assets - the 1. .py bundle, 2. bat files, and 3. two 601 and 602 mrs and dms scripts.

You need python and some basic dependencies (like pandas... - nothing outstanding) to run it, as well as Unicome environment (installed by default on our machines).

There are 3 BAT files - one for "load definitions", one for "apply definitions" and one for calling 601 mrs and 602 dms. Same as in older AA.

Please note BAT files do not have `pause` at the end - I just don't like it. If the command prompt disappeared - probably the file finished executing successfully.

Open BAT files and edit path to your MDD and map.

If you want to derive from older AA, make sure older AA is applied on MDD first. The, definitions will be pulled into new map.

Run `run_aarevival_reload_map.bat`. The map is created.

Send the map to team and wait for updates.

Once the map is updated, run `run_aarevival_apply_write_mrs.bat`. A mrs script `601_SavPrepRevival.INCLUDE.mrs` will be generated.

Run `run_aarevival_build_spss.bat`.

## How to verify that the new SPSS is close to the previous created with AA? ##
Use this tool: 
[Diff](https://github.com/andreyputilovmaterial/MDM-Diff-py/)

## Key featres ##
It follows basically the same usage pattern as older AA.

A map is in Excel. It no longer has any macros.

All validation is done with formulas and you don't have to wait for it - inputs are validated within second. Validation inlcudes information on items that are wrong (questions and categories), highlight problematic items with red or missing with yellow. It's easy to use.

Everything should look super simple for project teams.

The critical issue is probably that I don't have autofill. This will be added later.

We are using same dms scripts from Unicom as always.

These scripts normally support Unicode, no problem with it.

These scripts do not have any problems with big files in trackers - it can handle big file sizes. SPSS generating is one of the fastest scripts in our workflow, other scripts take more.

As of current implementation, the script does not write definitions to MDD, it does not modify the MDD, as opposed to older AA. I know, some people were hating this behavior. Now the script does not modify your MDDs.

As of current implementation, the script generates a piece of mrs code that is included in 601_SavPrep.mrs. Maybe this step can be moved to python as well - why are generating piece of mrs code that produces S-MDD when we can just create S-MDD instead?

Also, the scripts do not have any validation. All validation is in Excel and it's just informative. Yeah, the page will be red, saying that some items are bad... But if we run the script, it runs. The reasons are: 1. I think having that validation in Excel is enough to control quality, 2. if some definitions are incompatible, dms scripts will break with proper message. For example, if a shortname is used twice, there is an error "Alias already used", if the name is not a valid name there is also an error... 3. Also, I said I don't have any validation in my scripts - in fact I have, the scripts will stop if definitions for some of the variables or categories are missing in the map - this is considered to be a critical issue.

All these scripts can be added to your workflow programmatically. It no longer requires pressing any button in Excel manually, all steps can be called programmatically.

Scripts do not have to be in one folder. That .py bundle should probably be moved somewhere to ./Includes, or to some other subfolder.

## Runtimes ##
Runtimes are super fast. Validation takes 1 second. Map reload within 1-2 minutes too. 601_SavPrep.mrs takes couple minutes to run. Then, 602_SavCreate.dms takes \~40-90 minutes to produce a 5 GB SPSS.

## Why this was created ##
As I see, in trackers with big size of data, Flatout can't handle generating SPSS and can't allocate enough memory and process the file.
Our old scripts are working absolutely fine. The only pain point is validating AA, which is not always clear and takes long to apply definitions.
By "not always clear" - I mean shortnames in loops and arrays. Some rows in AA really mean NOTHING. Here, in new AA, every row is a variable. I tried to create everything nice and clear.

So, I see, the problem is to wait when AA is applied, validated, etc... Actually, I realized, we just make this new updated version of AA, and we no longer have any problems. That's it. Simple.

Another issue is that files produced with dms scripts are often not in Unicode but that's not a problem at all. Unicode is normally supported from 7.5.1 release - see release notes.

My goal is not to replace any other tools, Flatout or anything. Just to provide a solution for something that is generally working, that was working for years, and to address that step with AA validation. The goal is "just for fun".

