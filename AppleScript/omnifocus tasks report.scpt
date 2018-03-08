on run
	set theNoteName to "OmniFocus 已完成任务报告"
	-- 弹窗让用户选择统计时间，目前提供 "今天", "昨天", "本周", "上周", "本月", "上月" 选择项目
	activate
	set theReportScope to choose from list {"今天", "昨天", "本周", "上周", "本月", "上月"} default items {"本周"} with prompt "生成报告:" with title theNoteName
	if theReportScope = false then return
	set theReportScope to item 1 of theReportScope
	
	
	-- 计算时间差值
	set theStartDate to current date

	-- 取本日零时时间戳
	set hours of theStartDate to 0
	set minutes of theStartDate to 0
	set seconds of theStartDate to 0
	 
	-- 默认结束时间为本日 23:59:59
	set theEndDate to theStartDate + (23 * hours) + (59 * minutes) + 59
	 
	-- 按照用户选择选取不同的时间范围
	if theReportScope = "今天" then
		set theDateRange to date string of theStartDate
	else if theReportScope = "昨天" then
		set theStartDate to theStartDate - 1 * days
		set theEndDate to theEndDate - 1 * days
		set theDateRange to date string of theStartDate
	else if theReportScope = "本周" then
	-- 一直往前找，直到周一
		repeat until (weekday of theStartDate) = monday
			set theStartDate to theStartDate - 1 * days
		end repeat
	-- 一直往后找，直到周日
		repeat until (weekday of theEndDate) = sunday
			set theEndDate to theEndDate + 1 * days
		end repeat
		set theDateRange to (date string of theStartDate) & " 至 " & (date string of theEndDate)
	else if theReportScope = "上周" then
	-- 先将准标提前到上周
		set theStartDate to theStartDate - 7 * days
		set theEndDate to theEndDate - 7 * days
	-- 紧接着和生成本周类似，生成上周的两个准标
		repeat until (weekday of theStartDate) = monday
			set theStartDate to theStartDate - 1 * days
		end repeat
		repeat until (weekday of theEndDate) = sunday
			set theEndDate to theEndDate + 1 * days
		end repeat
		set theDateRange to (date string of theStartDate) & " 至 " & (date string of theEndDate)
	else if theReportScope = "本月" then
	-- 初始日期调整到本月 1 日
		repeat until (day of theStartDate) = 1
	 				set theStartDate to theStartDate - 1 * days
		end repeat
	-- 将结束日期调整到本月最后一天
		repeat until (month of theEndDate) is not equal to (month of theStartDate)
			set theEndDate to theEndDate + 1 * days
		end repeat
	-- 已经到下一个月的第一天了，往回倒退一天
		set theEndDate to theEndDate - 1 * days
		set theDateRange to (date string of theStartDate) & " 至 " & (date string of theEndDate)
	else if theReportScope = "上月" then
	-- 如果 1 月份，需要调整到上一年
		if (month of theStartDate) = January then
			set (year of theStartDate) to (year of theStartDate) - 1
			set (month of theStartDate) to December
		else
			set (month of theStartDate) to (month of theStartDate) - 1
		end if
	-- 调整到正确的年份和月份之后，保持和上面本月周报获取锚点一致
		set month of theEndDate to month of theStartDate
		set year of theEndDate to year of theStartDate
		repeat until (day of theStartDate) = 1
			set theStartDate to theStartDate - 1 * days
		end repeat
		repeat until (month of theEndDate) is not equal to (month of theStartDate)
			set theEndDate to theEndDate + 1 * days
		end repeat
		set theEndDate to theEndDate - 1 * days
		set theDateRange to (date string of theStartDate) & " 至 " & (date string of theEndDate)
	end if

	set ignoreList to {""}
	set reportName to theReportScope & "工作报告" & ".md"

	set theProgressDetail to "# 已完成任务 - " & theDateRange & return
	
	
	tell application "OmniFocus"
		-- 添加标识用来确保不会产生空的报表
		set modifiedTasksDetected to false

		tell default document
			set theModifiedProjects to every flattened project where its modification date is greater than theStartDate
			-- 遍历所有的 Project
			repeat with a from 1 to length of theModifiedProjects
				set theCurrentProject to item a of theModifiedProjects

				-- 取出指定 Project 中所有的已完成 Task 列表
				set theCompletedTasks to (every flattened task of theCurrentProject where its completed = true and completion date is greater than theStartDate and completion date is less than theEndDate and number of tasks = 0)
	 
				-- 循环每一个 Task
				if theCompletedTasks is not equal to {} then
					set modifiedTasksDetected to true
					-- Append the project name to the task list
					set theProgressDetail to theProgressDetail & "## " & name of theCurrentProject & return
	 
					repeat with b from 1 to length of theCompletedTasks
						set theCurrentTask to item b of theCompletedTasks
	 
						-- 目前只追加 Task 的 name
						set theProgressDetail to theProgressDetail & b & ". [" & name of theCurrentTask & "](" & note of theCurrentTask & ")" & return
					end repeat
	 
					set theProgressDetail to theProgressDetail & return
				end if
			end repeat

		end tell		
		 
		-- 将报表内容写入文件，提示用户下载到哪里
		set theProgressDetail to theProgressDetail as Unicode text
		set fn to choose file name with prompt "文件存储到：" default name reportName default location (path to desktop folder)
		 
		tell application "System Events"
			set fid to open for access fn with write permission
			write theProgressDetail to fid as «class utf8»
			close access fid
		end tell	
	end tell	
end run
