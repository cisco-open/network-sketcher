# Network Sketcher　Ver 2.1.0
# Update
* VPN Diagram
* Cross-platform support (Windows, Linux, Mac OS)
* Support for drawing beyond maximum PowerPoint size
* Import of CML diagrams (migration from ver1.x)

# Known Bug
| Known Bug                                                                                                               | Workaround|
|-------------------------------------------------------------------------------------------------------------------------| ------------- |
| [ns-bug-002] If the "Way Points" between the "Areas" are not interconnected, the generation of "Master Data" will fail. | Move unconnected "Way Point" between "Area" up/down the "Area".  |
| [ns-bug-003] On Mac OS, file cannot be opened from the dialog after drawing a diagram  | Click to open generated file  |


# Resolved
| Resolved Bug                                                                                                                                                | Workaround                                                      | Resolved Version|
|-------------------------------------------------------------------------------------------------------------------------------------------------------------|-----------------------------------------------------------------|------------- |
| [ns-bug-001] If a "way point" that does not exist above or below the leftmost "Area" exists in the right "Area", the generation of "Master Data" will fail. | Move unconnected "Way Point" between "Area" up/down the "Area". | 1.12|
