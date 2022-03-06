import re


class Solution:
    def isMatch(self, s: str, p: str, isMatch_flag: bool) -> bool:
        # isMatch_flag代表忽略条件，不参与匹配，从而始终返回True
        if isMatch_flag == False:
            return True
        else:
            lp, ls = len(p), len(s)
            dp = [set() for _ in range(lp + 1)]
            dp[0].add(-1)
            for i in range(lp):
                if p[i] == '?':
                    for x in dp[i]:
                        if x + 1 < ls:
                            dp[i + 1].add(x + 1)
                elif p[i] == '*':
                    minx = ls
                    for x in dp[i]: minx = min(minx, x)
                    while minx < ls:
                        dp[i + 1].add(minx)
                        minx += 1
                else:
                    for x in dp[i]:
                        if x + 1 < ls and s[x + 1] == p[i]:
                            dp[i + 1].add(x + 1)
                        elif x + 1 < ls and s[i] in ["'", '"'] and p[i] in ["'", '"']:
                            dp[i + 1].add(x + 1)
            if ls - 1 in dp[-1]:
                return True
            # 注释，改为单向匹配
            # if "*" in s or "*" in p:
            #     result = self.isMatch2(s=p,p=s)
            #     return result
            return False

    def isMatch2(self, s: str, p: str) -> bool:
        lp, ls = len(p), len(s)
        dp = [set() for _ in range(lp + 1)]
        dp[0].add(-1)
        for i in range(lp):
            if p[i] == '?':
                for x in dp[i]:
                    if x + 1 < ls:
                        dp[i + 1].add(x + 1)
            elif p[i] == '*':
                minx = ls
                for x in dp[i]: minx = min(minx, x)
                while minx < ls:
                    dp[i + 1].add(minx)
                    minx += 1
            else:
                for x in dp[i]:
                    if x + 1 < ls and s[x + 1] == p[i]:
                        dp[i + 1].add(x + 1)
        if ls - 1 in dp[-1]:
            return True
        return False

    def isMatch_compile(self, s: str, p: str) -> bool:
        """
        s:source str
        p:compile str
        Return bool, True if match
        """
        pattern = re.compile(p)
        if pattern.match(s):
            return True
        else:
            return False
