tell application "OmniFocus"
	tell front document
		set today to current date
		set hours of today to 0
		set minutes of today to 0
		set seconds of today to 0
		
		set yesterdayOffset to -1
		try
			set dialogResult to display dialog "Enter 'yesterday' offset:" default answer yesterdayOffset
			set yesterdayOffset to text returned of dialogResult as number
		end try
		
		set yesterday to today + yesterdayOffset * days
		set tomorrow to today + 1 * days
		
		set resultText to "## 어제 한 일"
		
		set allProjects to every flattened project
		repeat with projectIndex from 1 to length of allProjects
			set currentProject to item projectIndex of allProjects
			
			set includeCurrentProject to false
			set projectTaskDetails to ""
			
			set projectTasks to every task of currentProject
			repeat with taskIndex from 1 to length of projectTasks
				set currentTask to item taskIndex of projectTasks
				
				set isWorkTask to false
				set taskTags to tags of currentTask
				repeat with taskTagIndex from 1 to length of taskTags
					set taskTag to item taskTagIndex of taskTags
					if (name of taskTag) is "Work" then
						set isWorkTask to true
					end if
				end repeat
				
				if isWorkTask = true then
					set includeCurrentTask to false
					set currentTaskCompletionDate to completion date of currentTask
					if currentTaskCompletionDate ≥ yesterday and currentTaskCompletionDate < tomorrow then
						set includeCurrentTask to true
					end if
					
					set subtaskDetails to ""
					set subtasks to (every task of currentTask where its completion date ≥ yesterday and completion date < tomorrow)
					repeat with subtaskIndex from 1 to length of subtasks
						set currentSubtask to item subtaskIndex of subtasks
						
						set includeCurrentTask to true
						set currentSubtaskName to name of currentSubtask
						
						if subtaskDetails is not "" then
							set subtaskDetails to subtaskDetails & return
						end if
						
						set subtaskDetails to subtaskDetails & "    - " & currentSubtaskName
					end repeat
					
					if includeCurrentTask = true then
						set includeCurrentProject to true
						set currentTaskName to name of currentTask
						set projectTaskDetails to projectTaskDetails & return & "* " & currentTaskName
						
						if subtaskDetails is not "" then
							set projectTaskDetails to projectTaskDetails & return & subtaskDetails
						end if
					end if
				end if
			end repeat
			
			if includeCurrentProject = true then
				set currentProjectName to name of currentProject
				set resultText to resultText & return & return & currentProjectName & return & projectTaskDetails
			end if
		end repeat
		
		set resultText to resultText & return & return & "## 오늘 할 일"
		
		set allProjects to every flattened project
		repeat with projectIndex from 1 to length of allProjects
			set currentProject to item projectIndex of allProjects
			
			set includeCurrentProject to false
			set projectTaskDetails to ""
			
			set projectTasks to every task of currentProject
			repeat with taskIndex from 1 to length of projectTasks
				set currentTask to item taskIndex of projectTasks
				
				set isWorkTask to false
				set taskTags to tags of currentTask
				repeat with taskTagIndex from 1 to length of taskTags
					set taskTag to item taskTagIndex of taskTags
					if (name of taskTag) is "Work" then
						set isWorkTask to true
					end if
				end repeat
				
				if isWorkTask = true then
					set includeCurrentTask to false
					set currentTaskCompleted to completed of currentTask
					set currentTaskDueDate to due date of currentTask
					if currentTaskCompleted is false and currentTaskDueDate ≥ today and currentTaskDueDate < tomorrow then
						set includeCurrentTask to true
					end if
					
					set subtaskDetails to ""
					set subtasks to (every task of currentTask where its completed is false and due date ≥ today and due date < tomorrow)
					repeat with subtaskIndex from 1 to length of subtasks
						set currentSubtask to item subtaskIndex of subtasks
						
						set includeCurrentTask to true
						set currentSubtaskName to name of currentSubtask
						
						if subtaskDetails is not "" then
							set subtaskDetails to subtaskDetails & return
						end if
						
						set subtaskDetails to subtaskDetails & "    - " & currentSubtaskName
					end repeat
					
					if includeCurrentTask = true then
						set includeCurrentProject to true
						set currentTaskName to name of currentTask
						set projectTaskDetails to projectTaskDetails & return & "* " & currentTaskName
						
						if subtaskDetails is not "" then
							set projectTaskDetails to projectTaskDetails & return & subtaskDetails
						end if
					end if
				end if
			end repeat
			
			if includeCurrentProject = true then
				set currentProjectName to name of currentProject
				set resultText to resultText & return & return & currentProjectName & return & projectTaskDetails
			end if
		end repeat
		
		display alert resultText
		
		set the clipboard to resultText
	end tell
end tell
