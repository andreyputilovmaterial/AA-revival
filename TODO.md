TODOs:

features
- Can derive from AA (AA needs to be applied to MDD, and then the new map should be loaded)
It is not 100% identical to old AA but we can walk through differences
basically, same method of generating SPSS is used. It means, it will be very similar
- Or, if we need to derive from FlatOut map, this could be possible in the future (haven't looked into it, but why would we need it?)
- Unicode is normally supported, this is not an issue (see release notes of 7.5.1 Unicom software)
- It processes files as our dms scripts, it means it is not stuck on big files. We need 5 GB SPSS, of 10 GB spss, or 20 GB SPSS - we just need to wait more (maybe 40 minutes, or 2 hours)
- No manual clicks are needed in this version of AA - we want to streamline the process with scripts - we can
- No validation! (explained above)
- The map has number of benefits:
- - Visually appealing validation (colored in red/blue as in old AA, and cell address is listed on validation tab)
- - No need to wait (all is updated with formulas, immediately)
- - Downside - requires relatively new Excel (at least from 2019 or later, haven't tested well though; might not work in open office; I mean, validation might not work; if the map is saved, we should be good. Content in cells is important; Formulas are not important; If formulas are not working well in some verison of Office Application - that's not a blocker)
- - Only real variables are included - no rows for outer grid or loop objects that do not mean actually a variable that were included in old AA and in Flatout, and these rows could look quite confusing)
- - As syntax sugar - we also have a longer description on variable type
- - There is a field for user comments, and comments are saved when the map is reloaded! Whoa!
- - For now, I don't have autofill. It means, if a new value is appears, it needs to be filled. Same as in older AA. Autofill is probably a necessary option when a new map is laoded - 
- filling it manually - project teams do not like it.
- - No requirement to have map named exactly as per MDD or to be located in the same folder.
- This new AA is in python but it uses our older dms scrips. Actually, bigger part can be transformed to Python. Why Python? Because it has everything we need, it is widespread. Simple and well-known approaches. Anyway we need some language, we can't implement all in vbscript and dms, those languages are outdated and very limited. I think Python is a good choice. I tried to have very little dependencies on external libraries. Openpyxl package is needed for formulas and formatting, Pandas package is needed as base package for working with Excel. Unicome tools are needed for reading MDD (every SE has it installed).


todo: autofill
todo: fill from top to bottom when validating analysis values
todo: rearrange scripts
todo: do not call apply_write_mrs.bat from build_spss.bat, please add a call to --program write_mrs instead

